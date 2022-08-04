# Interview AddIn

Write a SOLIDWORKS add-in that creates a macro feature called BoundingBox. The macro feature calculates the bounding box dimensions of the document (sldprt and sldasm).

# Requirements

Do not use the built-in bounding box feature.

The dimensions should be called: Width, Length and Height. Height should be called Thickness if the document (part document) is a sheet metal one. 

The bounding box macro feature should be editable from a PMP (You must use the PMP built-in controls. ActiveX and .NET controls are not allowed. 
 ) and has the following options: 

- [ ] Tighest fit. Use [Ibody2::GetExtremePoints](https://help.solidworks.com/2021/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ibody2~getextremepoint.html) to get the dimensions. The option must always be checked. 
- [ ] Recalculates upon rebuild?
- [ ] Recalculates before closing (destroying)?
- [ ] Use document unit of length?

# Framework specs:

- SW 2019 or newer
- Code must be written in C#, .NET 4.6.2
- Minimal use of third-party dependencies. No SOLIDWORKS frameworks (SDK, Xarial...) are allowed.


# Plus points
It would be nice to see a history of commits instead of a blob of code. This allows us to see how you built the add-in.


# How to push your code 

Clone repo

Create a new branch with your name

Do a pull request 

# Conditions and limitations

Your code will be made available publicly under the MIT license. Please let me know if this is going to be an issue. 
