run("Clear Results"); // clear the results table of any previous measurements
//Needs SEM cross-section images
//Semi-Automated pore sizer where particles drawn by hand 
//final entry in csv being the particle in which the pores were identified


inputDirectory = getDirectory("Choose a Directory of Images");
fileList = getFileList(inputDirectory);
outputDirectory = getDirectory("Choose an output Directory");


for (i = 0; i < fileList.length; i++)
{
    //Cuting out the scale bar and detail footer
    //Need to update scalebar based on image acquired 
    open(inputDirectory+fileList[i]);
    run("Set Scale...", "distance=221 known=20 unit=Micron global");
    run("Set Measurements...", "area center bounding fit shape redirect=None decimal=3");
    //setTool("rectangle");
    makeRectangle(0, 0, 1536, 1024); //Thermo Axia typical image size
    run("Crop");
    setOption("BlackBackground", true);
    run("Convert to Mask");
    run("Invert");
	run("Despeckle");
	
    //Particle Identifier-Manually drawing ROI inside which we can process the pores
    PN = getNumber("How many particles in this image?",2);
    for (j = 0; j < PN; j++) 
    {
         setTool("freehand");
         waitForUser("Waiting for user to draw around particle");
         roiManager("Add");
         roiManager("Select", j);
         roiManager("Show All");
         
         //Pore size and shape factors analyzed first
         run("Set Measurements...", "area center perimeter bounding fit shape feret's redirect=None decimal=3");
         run("Analyze Particles...", "display clear include");
         run("Analyze Particles...", "display exclude clear include");
         //Bounding Particle detail will be added to pore detail as the last entry
         roiManager("Select", j);
         run("Measure");
         
         //Saving results
         selectWindow("Results");
         saveAs("Results", outputDirectory+"Image"+i+1+"Particle"+j+1+".csv");
         
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

close("*");

if (isOpen("Results")) 
  {
    selectWindow("Results");
    run("Close");
  }      
     







