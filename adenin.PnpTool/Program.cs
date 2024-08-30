using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using PnP.Core.Auth.Services.Builder.Configuration;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using PnP.Core.Services.Builder.Configuration;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;

Console.WriteLine("Would you like to provision Viva Dashboard [V] or a SharePoint Page [S]? V/S:");

var operation = Console.ReadLine()!.ToUpper();

if (operation == "S")
{
    Console.WriteLine("What's the URL of the SharePoint site to which you want to add 'coe.aspx', for example\r\n{your-sharepoint-domain}/sites/leadership:");
}
else
{
    Console.WriteLine("What's your SharePoint's base domain:");
}

var sharepointSite = Console.ReadLine()!;

var host = Host
    .CreateDefaultBuilder()
    .ConfigureServices((_, services) =>
    {
        services.AddPnPCore(options =>
        {
            options.PnPContext.GraphFirst = true;
            options.HttpRequests.UserAgent = "ISV|Contoso|ProductX";

            options.Sites.Add("SiteToWorkWith", new PnPCoreSiteOptions
            {
                SiteUrl = sharepointSite
            });
        });

        services.AddPnPCoreAuthentication(options =>
        {
            options.Credentials.Configurations.Add("interactive", new PnPCoreAuthenticationCredentialConfigurationOptions
            {
                ClientId = "7664b1d2-0f3c-47e2-bc78-a2703a2cfd7b",
                Interactive = new PnPCoreAuthenticationInteractiveOptions
                {
                    RedirectUri = new Uri("http://localhost")
                }
            });
            
            options.Credentials.DefaultConfiguration = "interactive";
            
            options.Sites.Add("SiteToWorkWith", new PnPCoreAuthenticationSiteOptions
            {
                AuthenticationProviderName = "interactive"
            });
        });
    })
    .UseConsoleLifetime()
    .Build();

await host.StartAsync();

const string appId = "744fb9bd-55d9-4bf3-8d22-3c9ee99819a5";
const string platformApi = "https://app.adenin.com/api/notebook/copy/templates-copilot/{0}";
const string cardUrlBase = "https://app.adenin.com/app/assistant/card/{0}";
const string apiErrorMessage = "Skipping {Card} as platform returned status {Status} during creation";
const string cardExistsWarning = "Skipping {Card} as it exists already in the dashboard";
const string cardUrlProp = "cardUrl";
const string cardUrlPlaceholder = "{0}";

const string vivaCardJson = $$"""
                          {
                              "appIntegrations": true,
                              "channel": "app_integrations",
                              "types": {
                                  "info": {
                                      "icon" :"ⓘ",
                                      "color": "emphasis"
                                  },
                                  "success": {
                                      "icon": "✅",
                                      "color": "good"
                                  },
                                  "error": {
                                      "icon": "❌",
                                      "color": "attention"
                                  },
                                  "blocked": {
                                      "icon": "🚫",
                                      "color": "attention"
                                  },
                                  "warning": {
                                      "icon": "⚠",
                                      "color": "warning"
                                  },
                                  "severeWarning": {
                                      "icon": "❗",
                                      "color": "warning"
                                  }
                              },
                              "cardUrl": "{{cardUrlPlaceholder}}"
                          }
                          """;

const string pageCardJson = $$"""
                          {
                              "description": "App integrations",
                              "height": "470px",
                              "borderToggle": true,
                              "appIntegrations": true,
                              "channel": "app_integrations",
                              "cardUrl": "{{cardUrlPlaceholder}}"
                          }
                          """;

var cardsJson = await File.ReadAllTextAsync("./cards.json");
var cardList = JsonSerializer.Deserialize<CardList>(cardsJson, SerializerOptions.CamelCase);

if (cardList is null)
{
    Console.Write("Could not find the card list JSON, cannot continue");
    return;
}

using (var scope = host.Services.CreateScope())
{
    var pnpContextFactory = scope.ServiceProvider.GetRequiredService<IPnPContextFactory>();
    var logger = scope.ServiceProvider.GetRequiredService<ILogger<Program>>();

    using var context = await pnpContextFactory.CreateAsync("SiteToWorkWith");

    var appManager = context.GetTenantAppManager();
    var integrationsApp = await appManager.AddAsync("./adenin-app-integrations.sppkg", true);
    var deploySucceeded = await appManager.DeployAsync(integrationsApp.Id);

    if (!deploySucceeded)
    {
        logger.LogError("Failed to deploy sppkg");
        return;
    }

    var token = await context.AuthenticationProvider.GetAccessTokenAsync(new Uri("https://graph.microsoft.com"));

    using var client = new HttpClient();

    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", $"office-365:{token}");

    switch (operation)
    {
        case "V":
            var dashboard = await context.Web.GetVivaDashboardAsync();

            if (dashboard is null)
            {
                logger.LogError("Site {Site} does not have a Viva dashboard instance", sharepointSite);
                return;
            }

            logger.LogInformation("Viva dashboard is available with {Count} cards currently existing", dashboard.ACEs.Count);

            foreach (var card in cardList.Cards)
            {
                var response = await client.GetAsync(string.Format(platformApi, card.Name));

                if (!response.IsSuccessStatusCode)
                {
                    logger.LogError(apiErrorMessage, card.Name, response.StatusCode);
                    continue;
                }

                var result = await response.Content.ReadFromJsonAsync<PlatformResponse<NotebookCopyResult>>();

                if (result?.ErrorCode is not 0)
                {
                    logger.LogError(apiErrorMessage, card.Name, result?.ErrorCode);
                    continue;
                }

                var cardUrl = string.Format(cardUrlBase, result.Data.Id);
                var exists = dashboard.ACEs.Any(ace => ace.Id == appId && ace.JsonProperties.TryGetProperty(cardUrlProp, out var property) && property.GetString() == cardUrl);

                if (exists)
                {
                    logger.LogWarning(cardExistsWarning, card.Name);
                    continue;
                }

                var customAce = dashboard.NewACE(Guid.Parse(appId));

                customAce.Title = result.Data.Title;
                customAce.CardSize = card.Size;

                if (result.Data.Logo is not null)
                {
                    customAce.IconProperty = result.Data.Logo;
                }

                customAce.Properties = JsonSerializer.Deserialize<JsonElement>(vivaCardJson.Replace(cardUrlPlaceholder, cardUrl));

                dashboard.AddACE(customAce);
            }

            await dashboard.SaveAsync();

            logger.LogInformation("Your Copilot dashboard is ready! Go to {SharepointSite}/SitePages/Dashboard.aspx to view it now. Thanks for installing, contact adenin if you have any questions at all!", sharepointSite);
            break;
        case "S":
            var page = await context.Web.NewPageAsync(PageLayoutType.Article);

            page.PageTitle = "Copilot Centre of Excellence dashboard - powered by adenin";
            page.SetCustomPageHeader("https://a.storyblok.com/f/150016/5472x3648/3d61f92558/laptop_bg.jpg");

            var numRows = cardList.Cards.Length / 3;
            var remainder = cardList.Cards.Length % 3;

            if (remainder > 0)
            {
                numRows++;
            }

            for (var i = 0; i < numRows; i++)
            {
                page.AddSection(CanvasSectionTemplate.ThreeColumn, 1);
            }

            var availableComponents = page.AvailablePageComponents().Where(c => c.ComponentType == 1);
            var thirdPartyWebPartComponent = availableComponents.First(c => c.Name == "App integrations");

            var currentRow = 0;
            var currentColumn = 0;

            foreach (var card in cardList.Cards)
            {
                var response = await client.GetAsync(string.Format(platformApi, card.Name));

                if (!response.IsSuccessStatusCode)
                {
                    logger.LogError(apiErrorMessage, card.Name, response.StatusCode);
                    continue;
                }

                var result = await response.Content.ReadFromJsonAsync<PlatformResponse<NotebookCopyResult>>();

                if (result?.ErrorCode is not 0)
                {
                    logger.LogError(apiErrorMessage, card.Name, result?.ErrorCode);
                    continue;
                }

                var cardUrl = string.Format(cardUrlBase, result.Data.Id);
                var webPart = page.NewWebPart(thirdPartyWebPartComponent);

                webPart.Title = result.Data.Title;
                webPart.PropertiesJson = pageCardJson.Replace(cardUrlPlaceholder, cardUrl);

                page.AddControl(webPart, page.Sections[currentRow].Columns[currentColumn]);

                currentColumn++;

                if (currentColumn > 2)
                {
                    currentRow++;
                    currentColumn = 0;
                }
            }

            await page.SaveAsync("coe.aspx");

            logger.LogInformation("Your Copilot dashboard is ready! Go to {SharepointSite}/SitePages/coe.aspx to view it now. Thanks for installing, contact adenin if you have any questions at all!", sharepointSite);
            break;
        default:
            logger.LogWarning("Operation {Operation} is not valid, exiting", operation);
            break;
    }
}

host.Dispose();

internal record PlatformResponse<T>(int ErrorCode, T Data);

internal record NotebookCopyResult(Guid Id, string Title, string? Logo);

internal record CardList(Card[] Cards);

internal record Card(string Name, CardSize Size);

internal static class SerializerOptions
{
    internal static JsonSerializerOptions CamelCase = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };
}
