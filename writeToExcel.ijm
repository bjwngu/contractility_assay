/*
 * Macro template to process multiple images in a folder
*/

#@ File (label = "ROI directory", style = "directory") input
#@ File (label = "Image directory", style = "directory") input_image
#@ File (label = "Output directory", style = "directory") output
suffix = ".zip"

processFolder(input, input_image);

// function to scan folders/subfolders/files to find files with correct suffix
function processFolder(input, input_image) 
{
	list = getFileList(input);
	Array.print(list);
	list2 = getFileList(input_image);
	Array.print(list2)
	for (i = 0; i < list.length; i++) 
	{
		if(File.isDirectory(input + File.separator + list[i]))
			processFolder(input + File.separator + list[i]);
		if(endsWith(list[i], suffix))
			processFile(input, input_image, output, list[i], list2[i]);
	}
}


function processFile(input, input_image, output, file, image_file) 
{
	// open file 
	open(input_image + File.separator + image_file);
	print("Opening Image: " + input_image + File.separator + image_file);
	open(input + File.separator + file);
	print("Opening ROI: " + input + File.separator + file);
	
	// Do the processing here by adding your own code.
	roiManager("Measure");
	// Change location to save .xlsx file to
	run("Read and Write Excel", "file=[C:/Users/benja/Desktop/contractility_assay_ngu/test.xlsx]");
	run("Close");
	close("*");
	roiManager("Delete");
    close("Results"); 
    close("ROI Manager");
    // Save and close file 
    // save(output + File.separator + file); 
    // close();
   
	// Leave the print statements until things work, then remove them.
	// print("Processing: " + input + File.separator + file);
	// print("Saving to: " + output);
}
