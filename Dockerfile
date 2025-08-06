

# Use the official .NET SDK image as a base image
FROM mcr.microsoft.com/dotnet/sdk:6.0  AS base
# Install libgdiplus in the runtime image
# This is crucial because your application will run in this stage
RUN apt-get update && apt-get install -y libgdiplus
WORKDIR /app

FROM mcr.microsoft.com/dotnet/sdk:6.0  AS build
# Copy the project files to the container
COPY . .

# Build the application
RUN dotnet restore
RUN dotnet publish -c Release -o /app/publish 



# Create the final runtime image
FROM mcr.microsoft.com/dotnet/aspnet:6.0
WORKDIR /app
COPY --from=build /app/publish .

# Set the entry point for the application
ENTRYPOINT ["dotnet", "BestPolicyReport.dll"]

# Create the Logs folder
RUN mkdir Logs

