FROM mcr.microsoft.com/dotnet/sdk:10.0.5-noble AS build
WORKDIR /src

COPY ["SistemaControlEspaciosDeportivosWeb.csproj", "./"]
RUN dotnet restore "SistemaControlEspaciosDeportivosWeb.csproj"

COPY . .
RUN dotnet publish "SistemaControlEspaciosDeportivosWeb.csproj" --configuration Release --output /app/publish /p:UseAppHost=false

FROM mcr.microsoft.com/dotnet/aspnet:10.0.5-noble AS final
WORKDIR /app

COPY --from=build --chown=$APP_UID:$APP_UID /app/publish .
USER $APP_UID

ENV ASPNETCORE_URLS=http://+:8080 \
    DataProtection__KeysPath=/app/data-protection

EXPOSE 8080
ENTRYPOINT ["dotnet", "SistemaControlEspaciosDeportivosWeb.dll"]
