# Use the official .NET SDK image as a base image for building
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


# Pre-configure ttf-mscorefonts-installer to accept the EULA in its own layer
RUN echo "debconf ttf-mscorefonts-installer/accepted-mscorefonts-eula select true" | debconf-set-selections

# Install libgdiplus and a comprehensive set of its common runtime dependencies.
# This now includes more common fonts and the command to refresh the font cache.
# Setting DEBIAN_FRONTEND to noninteractive ensures no prompts during installation.
RUN DEBIAN_FRONTEND=noninteractive 
RUN apt-get install -y  libgdiplus
RUN apt-get install -y  fontconfig 
RUN apt-get install -y  fonts-dejavu-core 
RUN apt-get install -y  ttf-mscorefonts-installer 
RUN apt-get install -y  fonts-liberation 
RUN apt-get install -y  libicu-dev 
RUN apt-get install -y  libcairo2-dev 
RUN apt-get install -y  libjpeg-dev 
RUN apt-get install -y  libpng-dev 
RUN apt-get install -y  libtiff-dev 
RUN apt-get install -y  libgif-dev 
RUN apt-get install -y  libxrender1 
RUN apt-get install -y  wget 
RUN apt-get install -y  ca-certificates
    
RUN rm -rf /var/lib/apt/lists/*

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