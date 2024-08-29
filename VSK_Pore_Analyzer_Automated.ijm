run("Clear Results"); // clear the results table of any previous measurements
//Needs SEM cross-section images
//Automated pore sizer where particles chosen by thresholding
//final entry in csv being the particle in which the pores were identified


inputDirectory = getDirectory("Choose a Directory of Images");
fileList = getFileList(inputDirectory);
outputDirectory = getDirectory("Choose an output Directory");


setBatchMode(true); 

//Particle Identifier-Provides ROI or particle spaces
for (i = 0; i < fileList.length; i++)
{
    //Cutting out the scale bar and detail footer
    //Need to update scalebar based on image acquired 
    open(inputDirectory+fileList[i]);
    run("Set Scale...", "distance=221 known=20 unit=Micron global");
    run("Set Measurements...", "area center bounding fit shape redirect=None decimal=3");
    //setTool("rectangle");
    makeRectangle(0, 0, 1536, 1024); 
    run("Crop");
    run("8-bit");
    setAutoThreshold("Default");
    //run("Threshold...");
    //setThreshold(0, 121);
    run("Convert to Mask");
    run("Despeckle");
    run("Despeckle");
    //run("Invert");
    //run("Fill Holes");
    run("Analyze Particles...", "size=3.5-Infinity display exclude clear add in_situ");
    close();
    
    
//Pore focused image pre-processing
    open(inputDirectory+fileList[i]);
    //setTool("rectangle");
    makeRectangle(0, 0, 1536, 1024);
    run("Crop");
    setOption("BlackBackground", true);
    run("Convert to Mask");

    
    PN = roiManager("Count");
    for (j = 0; j < PN; j++) 
    {	 
         roiManager("Select", j);
         run("Despeckle");
         
         //Pore size and shape factors analyzed first
         run("Set Measurements...", "area center perimeter bounding fit shape feret's redirect=None decimal=3");
         run("Analyze Particles...", "display clear include in_situ");
         
         //Bounding Particle detail will be added to pore detail as the last entry
         roiManager("Select", j);
         run("Measure");
         
         //Saving results
         selectWindow("Results");
         saveAs("Results", outputDirectory+"Image"+i+"P"+j+".csv");
         
         //Closing Results window for jth particle and pore detail analysis
         if (isOpen("Results")) 
          {
            selectWindow("Results");
            run("Close");
          }  
    }
    
     //Closing ROI manager at the end of processing ith image   
     if (isOpen("ROI Manager")) 
     {
         selectWindow("ROI Manager");
         run("Close");
     }      
     
     
    selectWindow(fileList[i]);
    close();
 
}

setBatchMode(false); 
close("*");

if (isOpen("Results")) 
  {
    selectWindow("Results");
    run("Close");
  }      
  