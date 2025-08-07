﻿# Use the official .NET SDK image as a base image for building
FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build

# Update apt and install build-time dependencies (if any)
# build-essential is added here as it provides common build tools that might be needed
# for native dependencies, though the primary issue is runtime.
RUN apt-get update && apt-get install -y apt-utils libc6-dev build-essential

WORKDIR /app

# Copy the project files to the container
COPY . .

# Build the application
RUN dotnet restore
RUN dotnet publish -c Release -o /app/publish

# Create the final runtime image
FROM mcr.microsoft.com/dotnet/aspnet:6.0

# Install libgdiplus and a comprehensive set of its common runtime dependencies.
# These packages cover font rendering, image formats (JPEG, PNG, GIF, TIFF),
# and other graphical necessities that System.Drawing and FastReport might use.
RUN apt-get update && apt-get install -y \
    libgdiplus \
    fontconfig \
    fonts-dejavu-core \
    libicu-dev \
    libcairo2-dev \
    ttf-mscorefonts-installer \
    fonts-liberation \
    libjpeg-dev \
    libpng-dev \
    libtiff-dev \
    libgif-dev \
    libxrender1 \ 
    && rm -rf /var/lib/apt/lists/*



# Update the font cache after installing new fonts

RUN fc-cache -f -v

WORKDIR /app
COPY --from=build /app/publish .

# Copy the wwwroot directory to the final runtime image
# This ensures static files like your FastReport templates are available
COPY BestPolicyReport/wwwroot ./wwwroot

# Set the entry point for the application
ENTRYPOINT ["dotnet", "BestPolicyReport.dll"]

# Create the Logs folder
RUN mkdir Logs