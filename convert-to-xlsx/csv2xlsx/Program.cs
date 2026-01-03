

static class Program
{
    public static async Task Main()
    {
        //await csv2xlsx.Convert.WriteOpenXmlAsync(args[0]);

        csv2xlsx.Demo.CreateHyperlinkSheet();
        csv2xlsx.Demo.CreateHyperlinkStyledSheet();
        csv2xlsx.Demo.CreateInternallinkSheet();
        csv2xlsx.Demo.CreateInternalLinkStyledSheet();
    }
}


