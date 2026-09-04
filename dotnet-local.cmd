@echo off
set "DOTNET_CLI_HOME=%~dp0.dotnet-cli-home"
set "DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1"
set "DOTNET_CLI_TELEMETRY_OPTOUT=1"
"%~dp0.dotnet\dotnet.exe" %*
