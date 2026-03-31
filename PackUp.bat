cd HomeworkViewer
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
xcopy D:\Administrator\Documents\GitHub\HomeworkViewer\bin\Debug\net10.0-windows\Resources D:\Administrator\Documents\GitHub\HomeworkViewer\bin\Release\net10.0-windows\win-x64\Resources /s /e /y
cd D:\Administrator\Documents\GitHub\HomeworkViewer\bin\Release\net10.0-windows\
ren D:\Administrator\Documents\GitHub\HomeworkViewer\bin\Release\net10.0-windows\win-x64 HomeworkViewer