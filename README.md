# Setup:
- If you are a VB.NET dev, you can install this [Html5 auto-completion provider VS.NET extension](https://github.com/VBAndCs/Vazor/blob/master/vbxmlCompletionProviderVSIX.zip?raw=true)
- If you are a VB.NET contributors, you can add these two files to your Roslyn fork, in the folder `Completion/CompletionProviders` in the project `Microsoft.CodeAnalysis.VisualBasic.Features`

Usage:
This completion provider will provide HTML5 auto completion in VB.NET XML Literals. 
It provide auto completion only inside vbxml tags:
```VB.NET
Dim x = <vbxml>
   <!—auto completion for HTML 5 is available here -->
</vbxml>
```


I did this for two reasons:
1. not to mess with other xml literals that doesn't deal with html.
2. to support [Vazor](https://github.com/VBAndCs/Vazor).


You can write `<%` and press `Ctrl+space` to get this block written for you:
```VB.NET
    <%= (Function()
              Return < />
           End Function)( )%>
``` 

where you can use conditions or other vb code to return an html node.

And you can write `<(` and press `Ctrl+space` to get this block written for you:
```VB.NET
     <%= (Iterator Function()
              For Each item In Collection
                   Yield <p><%= item %></p>
              Next
           End Function)( ) %>
``` 

where you can modify it to iterat through your collection and yiled an thml node based on each item in the collection, like filling a list with elements.

# Note:
My aim from publishing this Roslyn-dependant source, is to allow contributors to build upon to provide a generic XML literals completion provider, based on the xsd that user supplies in the context as this missing feature was [described in the docs](https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-studio-2013/bb531325%28v%3dvs.120%29#enabling-xml-intellisense-in-visual-basic)
