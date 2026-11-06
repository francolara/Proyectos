FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
WORKDIR /src
COPY GestionServicios.Domain/GestionServicios.Domain.csproj GestionServicios.Domain/
COPY GestionServicios.Contracts/GestionServicios.Contracts.csproj GestionServicios.Contracts/
COPY GestionServicios.Application/GestionServicios.Application.csproj GestionServicios.Application/
COPY GestionServicios.Infrastructure/GestionServicios.Infrastructure.csproj GestionServicios.Infrastructure/
COPY GestionServicios.Web/GestionServicios.Web.csproj GestionServicios.Web/
RUN dotnet restore GestionServicios.Web/GestionServicios.Web.csproj
COPY GestionServicios.Domain/ GestionServicios.Domain/
COPY GestionServicios.Contracts/ GestionServicios.Contracts/
COPY GestionServicios.Application/ GestionServicios.Application/
COPY GestionServicios.Infrastructure/ GestionServicios.Infrastructure/
COPY GestionServicios.Web/ GestionServicios.Web/
RUN dotnet publish GestionServicios.Web/GestionServicios.Web.csproj -c Release --no-restore -o /out /p:UseAppHost=false
FROM mcr.microsoft.com/dotnet/aspnet:10.0 AS runtime
WORKDIR /app
RUN mkdir -p /var/lib/hyt/keys && chown -R ${APP_UID}:${APP_UID} /var/lib/hyt
COPY --from=build /out/ ./
ENV ASPNETCORE_HTTP_PORTS=8080
EXPOSE 8080
USER $APP_UID
ENTRYPOINT ["dotnet", "GestionServicios.Web.dll"]
