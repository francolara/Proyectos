FROM mcr.microsoft.com/dotnet/sdk:10.0-noble@sha256:4ea6fe75dd36706bb6d8c3c293d4c4315840f5d76ea28ac97def77e3ec487fa5 AS build
WORKDIR /src

COPY ["SistemaControlEspaciosDeportivosWeb.csproj", "./"]
RUN dotnet restore "SistemaControlEspaciosDeportivosWeb.csproj"

COPY . .
RUN dotnet publish "SistemaControlEspaciosDeportivosWeb.csproj" --configuration Release --output /app/publish /p:UseAppHost=false

FROM mcr.microsoft.com/dotnet/aspnet:10.0-noble@sha256:1fe86375600b62e6566b465da9553eef0621f13c67f40fe764cd8dbb1dee1497 AS final
WORKDIR /app

COPY --from=build --chown=$APP_UID:$APP_UID /app/publish .
USER $APP_UID

ENV ASPNETCORE_URLS=http://+:8080 \
    DataProtection__KeysPath=/app/data-protection

EXPOSE 8080
ENTRYPOINT ["dotnet", "SistemaControlEspaciosDeportivosWeb.dll"]
