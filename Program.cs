using DeepLExcel;

if (args.Length < 6)
{
    Console.WriteLine("Usage: dotnet run <filename> <DeepL authkey> <skip header flag> <target column> <source language> <target language>");
}
else if (!File.Exists(args[0]))
{
    Console.WriteLine($"File {args[0]} doesn't exist");
}
else
{
    var translator = new ExcelTranslator(args[1], args[0]);
    bool skipHeader = false;
    bool.TryParse(args[2], out skipHeader);
    await translator.TranslateFile(skipHeader, args[3], args[4], args[5]);
}