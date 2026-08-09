FROM mcr.microsoft.com/dotnet/sdk:8.0-bookworm-slim AS restore
WORKDIR /src
COPY Directory.Build.props Directory.Packages.props pptx-mcp.sln ./
COPY src/PptxMcp.Api/PptxMcp.Api.csproj src/PptxMcp.Api/
COPY src/PptxMcp.Api/packages.lock.json src/PptxMcp.Api/
COPY tests/PptxMcp.Tests/PptxMcp.Tests.csproj tests/PptxMcp.Tests/
COPY tests/PptxMcp.Tests/packages.lock.json tests/PptxMcp.Tests/
RUN dotnet restore pptx-mcp.sln --locked-mode

FROM restore AS build
COPY . .
RUN dotnet build pptx-mcp.sln --configuration Release --no-restore

FROM node:20-bookworm-slim AS visual-renderer
WORKDIR /visual-renderer
COPY visual-renderer/package.json visual-renderer/package-lock.json ./
RUN npm ci --ignore-scripts --no-audit --no-fund
COPY visual-renderer/index.mjs visual-renderer/music-glyphs.mjs ./
COPY visual-renderer/assets ./assets
COPY visual-renderer/test ./test
RUN npm test

FROM build AS test
ENTRYPOINT ["dotnet", "test", "pptx-mcp.sln", "--configuration", "Release", "--no-build"]

FROM build AS publish
RUN dotnet publish src/PptxMcp.Api/PptxMcp.Api.csproj --configuration Release --no-build --output /app/publish

FROM mcr.microsoft.com/dotnet/aspnet:8.0-bookworm-slim AS runtime
ARG DEBIAN_FRONTEND=noninteractive
RUN apt-get update \
    && apt-get install --yes --no-install-recommends \
        curl \
        fonts-liberation \
        fonts-noto-cjk \
        libreoffice-impress \
        nodejs \
        poppler-utils \
    && rm -rf /var/lib/apt/lists/* \
    && mkdir -p /data/pptx-mcp /data/librechat-uploads \
    && chown -R app:app /data/pptx-mcp
WORKDIR /app
COPY --from=publish /app/publish .
COPY --from=visual-renderer --chown=app:app /visual-renderer /app/visual-renderer
ENV ASPNETCORE_HTTP_PORTS=8080 \
    DOTNET_EnableDiagnostics=0 \
    HOME=/tmp/app-home
USER app
EXPOSE 8080
ENTRYPOINT ["dotnet", "PptxMcp.Api.dll"]
