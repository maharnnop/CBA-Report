# Use the official .NET SDK image as a base image for building
FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build

# Update apt and install build-time dependencies (if any)
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
#RUN echo "debconf ttf-mscorefonts-installer/accepted-mscorefonts-eula select true" | debconf-set-selections

# Install libgdiplus and a comprehensive set of its common runtime dependencies.
# This now includes more common fonts and the command to refresh the font cache.
# Setting DEBIAN_FRONTEND to noninteractive ensures no prompts during installation.
# wget and ca-certificates are crucial for apt-get to fetch packages securely.
RUN DEBIAN_FRONTEND=noninteractive apt-get update && apt-get install -y \
    libgdiplus \
    fontconfig \
    fonts-dejavu-core \
    ttf-mscorefonts-installer \
    fonts-liberation \
    libicu-dev \
    libcairo2-dev \
    libjpeg-dev \
    libpng-dev \
    libtiff-dev \
    libgif-dev \
    libxrender1 \
    wget \
    ca-certificates \
    # Clean up apt cache to keep image size down
    && rm -rf /var/lib/apt/lists/*

# Update the font cache after installing new fonts
RUN fc-cache -f -v

WORKDIR /app
COPY --from=build /app/publish .

# Copy the wwwroot directory to the final runtime image
# This ensures static files like your FastReport templates are available
# Ensure 'BestPolicyReport/wwwroot' is the correct path relative to your Docker build context.
COPY BestPolicyReport/wwwroot ./wwwroot

# Set the entry point for the application
ENTRYPOINT ["dotnet", "BestPolicyReport.dll"]

# Create the Logs folder
RUN mkdir Logs
