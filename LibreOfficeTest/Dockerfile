# Stage 1: Build .NET app
FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build
WORKDIR /src
COPY . .
RUN dotnet restore
RUN dotnet publish -c Release -o /app/publish

# Stage 2: Runtime with LibreOffice + Python
FROM mcr.microsoft.com/dotnet/aspnet:6.0 AS final
WORKDIR /app

# Install LibreOffice Impress + headless + Python
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        fonts-dejavu \
        fonts-liberation \
        fonts-freefont-ttf \
        fonts-noto \
        fonts-roboto \
        fonts-font-awesome \
        fonts-opensymbol \
        fonts-wqy-microhei \
        fontconfig \
        libreoffice-impress \
        libreoffice-script-provider-python \
        python3 \
        python3-pip && \
    rm -rf /var/lib/apt/lists/* && \
    fc-cache -f

# Copy published .NET app
COPY --from=build /app/publish .

# Copy your Python scripts
COPY Python/ /app/python/

# Set ASP.NET Core to listen on 8080
ENV ASPNETCORE_URLS=http://+:8080
EXPOSE 8080

# Start the app
ENTRYPOINT ["dotnet", "LibreOfficeTest.dll"]
