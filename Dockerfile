

# Use the official .NET SDK image as a base image
FROM mcr.microsoft.com/dotnet/sdk:6.0  AS build
RUN apk add libgdiplus --repository http://dl-.alpinelinux.org/alpine/edge/testing/ 
RUN apk add ttf-freefont libssl1.1 
WORKDIR /app

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

