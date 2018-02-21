
# Release Notes Generator - Readme

Release Notes Template Replacement Tags
Anywhere the piece of data in the left column is found in your release notes, replace it with the tag from the column on the right, and the Release Notes Generator will replace the tag with the data collected by the script GUI. When using replacement tags to create a release notes template, it is important that the replacement tag has the same formatting as the text it is replacing. 

![alt text](http://bluegalaxy.info/images/rngt.jpg )
	

### Setup instructions
Decide where you want the Release Notes Generator to reside on your hard drive. The location where the .exe is unzipped will be its default ‘home’ directory. The script will look for a folder in this directory called “templates” where you will store the release notes templates. In addition to the script using the templates folder, it will create a “scratch” folder, and requires a “new_rn” folder. The scratch folder can be ignored as it is just used by the script. The “new_rn” folder is where the generated release notes will appear after running the script.


![alt text](http://bluegalaxy.info/images/rngf.png)
 
When running the script, a window will pop up where you choose the inputs that will be used as the basis for your release notes. Error messages will pop up if one of the inputs is not selected.


![alt text](http://bluegalaxy.info/images/rngi.png)
 
Note: The templates you have in your templates folder will determine the products available to choose from under “Select Product:”
	Templates should be named in two ways:
1.	Product name only for generic templates. i.e. “3D Landmarks.docx”
2.	Product name + underscore + region, for region specific templates. i.e. “3D Landmarks_WEU.docx”. 
Note: if you choose a region specific template under “Select Product:”, then the region you choose under “Select Region:” is irrelevant because the script will use the region specific to the template.

Click “Generate Release Notes” when selections have been made. The script will alert you when the release notes have been generated.


![alt text](http://bluegalaxy.info/images/rngm.png)
 

When first opening the newly created release notes, you will have to click through two Word error messages. For example:

![alt text](http://bluegalaxy.info/images/rnge.png)

Click OK.

![alt text](http://bluegalaxy.info/images/rnge2.png)
 
Click Yes.

Once opened, check each page to make sure all template replacements have been completed successfully. Then SAVE the file. Once saved, it can be re-opened without triggering the Word error message shown above.

