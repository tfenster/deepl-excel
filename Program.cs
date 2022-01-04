using DeepLExcel;

if (args.Length < 1)
{
    Console.WriteLine("Usage: dotnet run <filename>");
}
else if (!File.Exists(args[0]))
{
    Console.WriteLine($"File {args[0]} doesn't exist");
}
else
{
    var translator = new ExcelTranslator(args[1], args[0]);
    await translator.TranslateFile();
}