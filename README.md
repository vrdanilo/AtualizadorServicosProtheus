# AtualizadorServicosProtheus
Criação de um programa em C# para ajudar na atualização de binários do protheus e ajuste de alguns parâmetros nos appserver.ini dos serviços.

Para criar o executavel: dotnet publish -r win-x64 /p:PublishSingleFile=true /p:IncludeNativeLibrariesForSelfExtract=true --self-contained true