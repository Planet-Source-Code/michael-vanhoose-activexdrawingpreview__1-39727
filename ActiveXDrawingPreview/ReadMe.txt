The first on its kind psc.  Explorer Add-on. ActiveX thumbnail viewer.  Just Like viewing bitmaps in the left side.  
View AutoCAD drawings thru Explorer. 
It doen't have to be just AutoCad files, it can be anything you want.
Also, ActiveX Desktop. Please vote. see preview.


It took along time to figure this out.  Please vote.


I know this works for Window 2000 I don't know for the others.  It should


You need to do a little with JavaScript. It's Not to hard. dont worry.
this is how it works.

1.  compile to ActiveXDrawingPreview.ocx     sounds simple

2.  Register the ActiveXDrawingPreview.ocx   

3.  Search the Registry for the clsid            looks like this = {C261592A-FACF-4CDB-ADFB-B56B0070F450}

Search using the filename ACtiveXdrawingpreview.ocx

write down clsid.  you need this later.

also add to keys to your clsid
you don't have to add this. but it turn off a message box in explorer say not safe.

    

                                    <your clsid here>
HKEY_CLASSES_ROOT\CLSID\{C261592A-FACF-4CDB-ADFB-B56B0070F450}\Implemented Categories

{7DD95801-9882-11CF-9FA9-00AA006C42C4}
{7DD95802-9882-11CF-9FA9-00AA006C42C4}



4.  look at the folder.htt that I included.
    Explorer uses this file when exploring your files on your computer. for links and stuff like that.

    know open your folder.htt c:\WINNT\Web\folder.htt 
     Put the code below in the same place I did in my folder.htt
	(Copy your folder.htt first, just in case)

   know change the clsid on the </object to match your clsid without { & } on both sides.



			//Start of code for project
			else if (IsDwgFile(ext))
                        {
                            Preview.innerHTML = '<p>' +
 change this line>>>                               '<object  ID="DrawingPreview"  WIDTH=185 HEIGHT=140 class=DwgPreview classid="clsid:C261592A-FACF-4CDB-ADFB-B56B0070F450">' +
                             	
                                '</object>';
                      		DrawingPreview.YourDrawingFile = item.path;
				
                        }
			//end of code for project

	//start of code for project
	function IsDwgFile(ext) {
            var types = ",dwg,";
            var temp = ","+ext+",";
            return types.indexOf(temp) > -1;
        }
	//end of code for project

5. open explorer and click on .dwg files


Active Desktop
 
may do a different project then DrawingPreview

1. You need to have Active Desktop installed. 

2. You need to prepare a simple webpage for the object. Something like this: 
 
<html>
<body>
<object align=center classid=clsid:{C261592A-FACF-4CDB-ADFB-B56B0070F450}>
<center>
</center>
</object>
</html>
</body>

 
3. On an empty spot on the desktop, contextual click and "Properties" 
4. "Web" tab, select View my Active Destop as a web page. 
5. Click "New", then "No" in the pop up window. 
6. In the "Location" field, enter the path (or browse to it) to the HTML code 
    created from the above example. 
7. Hit apply. Now you have an object window floating without seams on the desktop. 



 
