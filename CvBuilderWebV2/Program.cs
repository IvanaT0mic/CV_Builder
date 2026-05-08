using CvBuilderWebV2;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using MudBlazor.Services;

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

builder.Services.AddMudServices();

var apiUrl = builder.Configuration["ApiUrl"];

builder.Services.AddScoped(sp => new HttpClient
{
    BaseAddress = new Uri(apiUrl!)
});

await builder.Build().RunAsync();