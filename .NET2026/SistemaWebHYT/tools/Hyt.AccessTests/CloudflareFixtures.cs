using System.Collections.Concurrent;
using GestionServicios.Application.Followup;
using GestionServicios.Web.Services;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.FileProviders;
sealed class StubCloudflareHandler:HttpMessageHandler
{
    public string Payload {get;set;}="{}";
    public System.Net.HttpStatusCode Status {get;set;}=System.Net.HttpStatusCode.OK;
    public bool Fail {get;set;}
    public int Requests {get;private set;}
    public Uri? Uri {get;private set;}
    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request,CancellationToken ct)
    {
        Requests++;Uri=request.RequestUri;if(Fail)throw new HttpRequestException("Simulación de red");
        return Task.FromResult(new HttpResponseMessage(Status){Content=new StringContent(Payload,System.Text.Encoding.UTF8,"application/json")});
    }
}
sealed class FakeR2Objects:IR2ObjectClient
{
    public ConcurrentDictionary<string,byte[]> Files {get;}=new();
    public Task<bool> PutIfAbsentAsync(DocumentAttachment metadata,byte[] bytes,CancellationToken ct)=>Task.FromResult(Files.TryAdd(metadata.ClaveAlmacenamiento,bytes.ToArray()));
    public Task<byte[]?> GetAsync(string key,CancellationToken ct)=>Task.FromResult(Files.TryGetValue(key,out var data)?data.ToArray():null);
}
sealed class TestHostEnvironment:IWebHostEnvironment
{
    public string EnvironmentName {get;set;}="Development";
    public string ApplicationName {get;set;}="Pruebas HYT";
    public string ContentRootPath {get;set;}=AppContext.BaseDirectory;
    public string WebRootPath {get;set;}=AppContext.BaseDirectory;
    public IFileProvider ContentRootFileProvider {get;set;}=new NullFileProvider();
    public IFileProvider WebRootFileProvider {get;set;}=new NullFileProvider();
}