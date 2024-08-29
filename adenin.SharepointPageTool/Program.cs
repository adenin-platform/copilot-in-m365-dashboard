using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using PnP.Core.Auth.Services.Builder.Configuration;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using PnP.Core.Services.Builder.Configuration;
using System.Buffers;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;

string sharepointSite;

if (args.Length == 0)
{
    Console.WriteLine("Please provide the Sharepoint site URL:");
    sharepointSite = Console.ReadLine()!;
}
else
{
    sharepointSite = args[0];
}

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

const string platformApi = "https://app.adenin.com/api/notebook/copy/templates-copilot/{0}";
const string cardUrlBase = "https://app.adenin.com/app/assistant/card/{0}";
const string apiErrorMessage = "Skipping {Card} as platform returned status {Status} during creation";
const string cardUrlPlaceholder = "{0}";
const string cardJson = $$"""
                          {
                              "description": "App integrations",
                              "height": "none",
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

    var pages = await context.Web.GetPagesAsync("coe.aspx");
    var page = pages.FirstOrDefault();

    if (page is not null)
    {
        await page.DeleteAsync();
    }

    page = await context.Web.NewPageAsync(PageLayoutType.Article);

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

    var token = await context.AuthenticationProvider.GetAccessTokenAsync(new Uri("https://graph.microsoft.com"));

    using var client = new HttpClient();

    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", $"office-365:{token}");

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
        webPart.PropertiesJson = cardJson.Replace(cardUrlPlaceholder, cardUrl);

        page.AddControl(webPart, page.Sections[currentRow].Columns[currentColumn]);

        currentColumn++;

        if (currentColumn > 2)
        {
            currentRow++;
            currentColumn = 0;
        }
    }

    await page.SaveAsync("coe.aspx");
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
