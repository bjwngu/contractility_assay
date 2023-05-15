//Benjamin Ngu 
//University of Southern California 
//Cedars-Sinai - Smidt Heart Institute - Parker Labratory
//Function: Per Cell Segmentation V3

// To add:
// - % saturation option (instead of just normalize) 

main();
setOption("ExpandableArrays", true);

// temp log 
var BPOper_temp=newArray; 
var BPIt_temp=newArray; 
var BPCt_temp=newArray; 
	
// perm log 
var images=newArray; 
var RBR_val=newArray; 
var EC=newArray; 
var thres=newArray; 
var BPOper=newArray;
var BPIt=newArray;
var BPCt=newArray; 

function main()
{
	startOption=openStartMenu();
	while (startOption)
	{
		waitForUser("Open New Image", "Press OK to continue"); 
		// Original Image Information 
		folder=getInfo("image.directory");
		filename=getInfo("image.filename");
		fullPath=folder+filename;	
		startSegmentation(filename, fullPath);
		startOption=openStartMenu();
	}

	// end program (summarize results)
	tabulate(); 
}

function startSegmentation(filename, fullPath)
{
	setOption("ExpandableArrays", true);

	EC_temp=' ';
	RBR_temp=' ';
	thres_temp=' '; 
		 
	status=true; 
	segState=0; 
	while (status)
	{
		state=progressionMenu(segState); // returns "restart" or "start/continue"
	 
		if (state=="Restart")
		{
			close("*");
			open(fullPath); 
			resetBPtemp();
			segState=0;    
		}
		else if (state=="Start/Continue")
		{
			if (segState==0) // ask full or cropped image? 
			{
				askCrop(); 
				segState++;  			
			}
			else if (segState==1) // 8bit + ask if enhance contrast? 
			{
				EC_temp=preProc(); 
				segState++;
			}
			else if (segState==2) // Background Subtraction
			{
				RBR_temp=subBackground(EC_temp); 
				segState++; 
			}
			else if (segState==3) //Threshold
			{
				thres_temp=thresh(); 
				segState++; 
			}
			else if (segState==4) // Binary Processing 
			{ 
				binProc();
				segState++; 
			}
			else if (segState==5) // Analyze Particles 
			{
				analyze(); 
				status=false; 
				open(fullPath); 
				writeToPerm(filename, EC_temp, RBR_temp, thres_temp);
				resetBPtemp(); 
			}
		}
	}
}

function analyze()
{
	run("Analyze Particles...", "size=10000-Infinity display clear include add");
	close(thres_temp); 
	close("Results"); 
}

function binProc()
{
	run("Fill Holes");
	Dialog.createNonBlocking("Start Binary Processing");
	Dialog.addHelp("https://docs.google.com/presentation/d/1_l-y74GNdpdKQDmrJSqrNwUxeJpR9bNM66gc-QqBmf4/edit?usp=sharing");
	Dialog.addMessage("Start binary processing if segmentation is not filled to completion");
	Dialog.addCheckbox("Check to start", true);
	Dialog.show();

	setOption("ExpandableArrays", true);
	if (Dialog.getCheckbox()==true)
	{
		continueBP=true;
		while(continueBP==true)
		{
			run("Undo");
			run("Options...");
			Dialog.createNonBlocking("Enter binary processing parameters");
			Dialog.addHelp("https://docs.google.com/presentation/d/1_l-y74GNdpdKQDmrJSqrNwUxeJpR9bNM66gc-QqBmf4/edit?usp=sharing");
			Dialog.addChoice("Operation (\"Do\")", newArray("Nothing", "Erode", "Dilate", "Open", "Close"));
			Dialog.addSlider("Iteration", 1, 20, 1);
			Dialog.addSlider("Count", 1, 3, 3);
			Dialog.show();

			BPOper_temp[BPOper_temp.length]=Dialog.getChoice();
			BPIt_temp[BPIt_temp.length]=Dialog.getNumber();
			BPCt_temp[BPCt_temp.length]=Dialog.getNumber(); 

			run("Fill Holes");
			Dialog.createNonBlocking("Contine binary processing if segmentation is not filled to completion");
			Dialog.addHelp("https://docs.google.com/presentation/d/1_l-y74GNdpdKQDmrJSqrNwUxeJpR9bNM66gc-QqBmf4/edit?usp=sharing");
			Dialog.addCheckbox("Check to continue BP", false);
			Dialog.show();
			continueBP=Dialog.getCheckbox();
		}
	}
	
	else 
	{
		BPOper_temp[BPOper_temp.length]="NA"; 
		BPIt_temp[BPIt_temp.length]="NA";
		BPCt_temp[BPCt_temp.length]="NA";
	}
}

function thresh()
{
	rename("MaxEntropy");
	run("Duplicate...", "title=Mean");
	run("Duplicate...", "title=Other");
	selectWindow("MaxEntropy");
	setAutoThreshold("MaxEntropy");
	setOption("BlackBackground", false);
	run("Convert to Mask");
	selectWindow("Mean");
	setAutoThreshold("Mean");
	setOption("BlackBackground", false);
	run("Convert to Mask");
	makeMontage(3);

	Dialog.createNonBlocking("Threshold");
	Dialog.addHelp("https://docs.google.com/presentation/d/1_l-y74GNdpdKQDmrJSqrNwUxeJpR9bNM66gc-QqBmf4/edit?usp=sharing");
 	algChoices = newArray("MaxEntropy", "Mean", "Other");
 	Dialog.addRadioButtonGroup("Choose thresholding algorithm", algChoices, 1, 3, algChoices[1]);
	Dialog.show();

	alg=Dialog.getRadioButton();
	if (alg=="Other")
	{
		selectWindow(alg);
		close("\\Others");
		run("Threshold...");
		Dialog.createNonBlocking("Select Other Algorithm");
		Dialog.addHelp("https://docs.google.com/presentation/d/1_l-y74GNdpdKQDmrJSqrNwUxeJpR9bNM66gc-QqBmf4/edit?usp=sharing");
		Dialog.addMessage("Choose and APPLY an alternative thresholding algorithm");
		Dialog.addString("Type selected algorithm", "Default");
		Dialog.show();
		alg=Dialog.getString();
		rename(alg);
		close("\\Others");
		close("Threshold");
	}
	else 
	{
		selectWindow(alg);
		close("\\Others");
	}

	return alg;
}

function subBackground(EC_temp)
{
	rename("Original");
	run("Duplicate...", "title=RBR=25");
	run("Duplicate...", "title=RBR=37.5");
	run("Duplicate...", "title=RBR=50");

	selectWindow("RBR=25");
	run("Subtract Background...", "rolling=25 light");
	selectWindow("RBR=37.5");
	run("Subtract Background...", "rolling=37.5 light");
	selectWindow("RBR=50");
	run("Subtract Background...", "rolling=50 light");
	makeMontage(4);

	Dialog.createNonBlocking("Background Subtraction");
	Dialog.addHelp("https://docs.google.com/presentation/d/1_l-y74GNdpdKQDmrJSqrNwUxeJpR9bNM66gc-QqBmf4/edit?usp=sharing");
 	RBRChoices = newArray("25", "37.5", "50", "Auto");
 	Dialog.addRadioButtonGroup("Choose a rolling ball radius", RBRChoices, 1, 3, "50");
	Dialog.show();

	RBR=Dialog.getRadioButton();
	if (RBR=="Auto")
	{
		selectWindow("Original");
		close("\\Others"); 
		
		run("8-bit");
		if (EC_temp=="Yes")
		{
			run("Enhance Contrast...", "saturated=0 equalize");
		}
		
		run("Subtract Background...");
		Dialog.createNonBlocking("Enter RBR");
		Dialog.addString("RBR", 25); 
		Dialog.show(); 
		RBR=Dialog.getString();
		rename("RBR=" + RBR); 
		close("\\Others"); 
	}
	else 
	{
		selectWindow("RBR=" + RBR);
		close("\\Others");
	}
	return RBR;
}

function preProc()
{
	run("8-bit");
	Dialog.createNonBlocking("Enhance contrast");
	Dialog.addHelp("https://docs.google.com/presentation/d/1_l-y74GNdpdKQDmrJSqrNwUxeJpR9bNM66gc-QqBmf4/edit?usp=sharing");
	Dialog.addRadioButtonGroup("Do you want to enhance image contrast?", newArray("Yes", "No"), 1, 2, "Yes");
	Dialog.show();

	yesOrNo=Dialog.getRadioButton();
	if (yesOrNo=="Yes")
	{
		// run("Enhance Contrast...", "saturated=0 equalize");
		run("Enhance Contrast...", "saturated=3 normalize");
	}

	return yesOrNo; 
}

function progressionMenu(segState)
{
	message=' '; 
	if (segState==0)
	{
		message="Proceed to: (0) Crop Image?"; 
	}
	else if (segState==1)
	{
		message="Proceed to: (1) 8-bit/Enhance Contrast?"; 
	}
	else if (segState==2)
	{
		message="Proceed to: (2) Background Subtraction"; 
	}
	else if (segState==3)
	{
		message="Proceed to: (3) Threshold";
	}
	else if (segState==4)
	{
		message="Procced to: (4) Binary Processing"; 
	}
	else if (segState==5)
	{
		message="Proceed to: (5) Analyze/Measure ROIs"; 
	}

	
	Dialog.createNonBlocking("Progression");
	Dialog.addHelp("https://docs.google.com/presentation/d/1_l-y74GNdpdKQDmrJSqrNwUxeJpR9bNM66gc-QqBmf4/edit?usp=sharing");
	progress=newArray("Start/Continue", "Restart"); 
	Dialog.addRadioButtonGroup(message, progress, 2, 1, progress[0]);
	Dialog.show();

	option=Dialog.getRadioButton();
	return option; 
}

function openStartMenu()
{
	option=true; 
	Dialog.createNonBlocking("Menu");
	Dialog.addHelp("https://docs.google.com/presentation/d/1_l-y74GNdpdKQDmrJSqrNwUxeJpR9bNM66gc-QqBmf4/edit?usp=sharing");
	Dialog.addRadioButtonGroup("Choose an option:", newArray("New Image", "End/Summarize"), 1, 2, "New Image");
	Dialog.show(); 
	if (Dialog.getRadioButton()=="End/Summarize")
	{
		option=false;
	}
	return option; 
}

function askCrop()
{
	Dialog.createNonBlocking("Crop Image");
	Dialog.addMessage("If segmenting a single cell, identify it with the rectangle tool");
	cropOrNot=newArray("Full", "Cropped");
	Dialog.addRadioButtonGroup("Do you want to crop your image?", cropOrNot, 1, 2, cropOrNot[1]);
	Dialog.addHelp("https://docs.google.com/presentation/d/1_l-y74GNdpdKQDmrJSqrNwUxeJpR9bNM66gc-QqBmf4/edit?usp=sharing");
	Dialog.show();
	if (Dialog.getRadioButton()=="Cropped")
	{
		 
		run("Duplicate...", " ");
		close("\\Others");
	}
}

function writeToPerm(filename, EC_temp, RBR_temp, thres_temp)
{ 
	 images[images.length]=filename; 
	 EC[EC.length]=EC_temp; 
	 RBR_val[RBR_val.length]=RBR_temp; 
	 thres[thres.length]=thres_temp; 

	 //currentIndex=BPOper.length; 
	 for (i=0; i<BPOper_temp.length; i++)
	 {
	 	BPOper[BPOper.length]=BPOper_temp[i];
	 	BPIt[BPIt.length]=BPIt_temp[i];  
	 	BPCt[BPCt.length]=BPCt_temp[i]; 
	 }

	 addBlanks=BPOper_temp.length-1;
	 for (i=0; i<addBlanks; i++)
	 {
	 	images[images.length]=' ';
	 	EC[EC.length]=' ';
	 	RBR_val[RBR_val.length]=' '; 
	 	thres[thres.length]=' '; 
	 }
}

function resetBPtemp()
{
	BPOper_temp=newArray; 
	BPIt_temp=newArray; 
	BPCt_temp=newArray;
}

function tabulate()
{
	Table.create("Summary of Parameters"); 
	Table.setColumn("Images", images);
	Table.setColumn("Enhance Contrast?", EC);
	Table.setColumn("RBR", RBR_val);
	Table.setColumn("Threshold", thres);
	Table.setColumn("Binary Operation", BPOper);
	Table.setColumn("Iterations", BPIt);
	Table.setColumn("Count", BPCt);
}

function makeMontage(numImages)
{
	numRows=1; 
	if (numImages==4)
	{
		numRows=2; 
	}
	run("Images to Stack", "name=myStack title=[] use keep");
	run("Make Montage...", "columns=" + numImages + " rows=1 scale=0.50 label");
	// run("Make Montage...", "Columns=" + numImages + "rows=" + numRows + " scale=2 label");
	close("myStack");
}


