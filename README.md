You can add these two files to the in your Roslyn fork, in the folder `Completion/CompletionProviders` in the project `Microsoft.CodeAnalysis.VisualBasic.Features`

It will provide HTML5 auto completion in VB.NET XML Literals, but only inside the `<vbxml></vbxml>` tag. 

I did this for two reasons:
1. not to mess with other xml literals that doesn't deal with html.
2. to support [Vazor](https://github.com/VBAndCs/Vazor).

I will make a VSIX for this to allow install it as a VS extension (in progress now), but my aim from publishing this, is to allow contributors to build upon to provide a generic XML literals completion provider, based on the xsd that user supplies in the context as this missing feature was [described in the docs](https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-studio-2013/bb531325%28v%3dvs.120%29#enabling-xml-intellisense-in-visual-basic)
