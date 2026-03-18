FROM mcr.microsoft.com/dotnet/aspnet:9.0 AS base
WORKDIR /app
EXPOSE 8080

# Python Embeddable for Windows (bundled in local version ZIP)
FROM python:3.12-slim AS python-prep
WORKDIR /prep
RUN apt-get update && apt-get install -y --no-install-recommends wget unzip \
    && rm -rf /var/lib/apt/lists/* \
    && wget -q https://www.python.org/ftp/python/3.12.8/python-3.12.8-embed-amd64.zip \
    && mkdir python-embed \
    && unzip python-3.12.8-embed-amd64.zip -d python-embed \
    && rm python-3.12.8-embed-amd64.zip
# Enable import site + add site-packages path
RUN sed -i 's/#import site/import site/' python-embed/python312._pth \
    && echo "Lib/site-packages" >> python-embed/python312._pth
# Download Windows-compatible wheel files
RUN pip download --platform win_amd64 --python-version 312 --only-binary :all: \
    gspread google-auth openpyxl -d /prep/wheels --no-cache-dir \
    && mkdir -p python-embed/Lib/site-packages \
    && for whl in /prep/wheels/*.whl; do unzip -o "$whl" -d python-embed/Lib/site-packages/; done \
    && rm -rf /prep/wheels

FROM mcr.microsoft.com/dotnet/sdk:9.0 AS build
WORKDIR /src
COPY ["ExcelMigrationApi.csproj", "."]
RUN dotnet restore
COPY . .
RUN dotnet publish -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=build /app/publish .
COPY --from=python-prep /prep/python-embed /app/python-embed
ENV ASPNETCORE_URLS=http://+:8080
ENTRYPOINT ["dotnet", "ExcelMigrationApi.dll"]
