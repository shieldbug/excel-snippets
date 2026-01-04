using System.Runtime.CompilerServices;

namespace Excel.Snippets
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            // await Convert.WriteOpenXmlAsync(args[0]);

            Demo.CreateHyperlinkSheet();
            Demo.CreateHyperlinkStyledSheet();
            Demo.CreateInternallinkSheet();
            Demo.CreateInternalLinkStyledSheet();
            Demo.CreateAutofitColumnSheet();

        }
    }
}
