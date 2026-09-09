FROM mcr.microsoft.com/dotnet/sdk:10.0.5-noble AS build
WORKDIR /src

COPY ["FralseTech.Web.csproj", "./"]
RUN dotnet restore "FralseTech.Web.csproj"

COPY . .
RUN dotnet publish "FralseTech.Web.csproj" --configuration Release --output /app/publish /p:UseAppHost=false

FROM mcr.microsoft.com/dotnet/aspnet:10.0.5-noble AS final
WORKDIR /app

COPY --from=build --chown=$APP_UID:$APP_UID /app/publish .
USER $APP_UID

ENV ASPNETCORE_URLS=http://+:8080

EXPOSE 8080
ENTRYPOINT ["dotnet", "FralseTech.Web.dll"]
