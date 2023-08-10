async function main(workbook: ExcelScript.Workbook) {

	const newSheet = workbook.addWorksheet("Final Form")


	let cellB31_ = workbook.getWorksheet("Final Form").getRange("B31");
	cellB31_.setValue("Signature: ");
	cellB31_.getFormat().getFont().setBold(true);
	cellB31_.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);

	let cellC31_ = workbook.getWorksheet("Final Form").getRange("C31");
	cellC31_.setValue("______________________________________");
	cellC31_.getFormat().getFont().setBold(true);

	cellB31_.getFormat().setRowHeight(40);

	// Prepare the formatting of all the cells:
	let cellA1_ = workbook.getWorksheet("Final Form").getRange("A1");
	let cellB1_ = workbook.getWorksheet("Final Form").getRange("B1");
	let cellC1_ = workbook.getWorksheet("Final Form").getRange("C1");

	let formatA1 = cellA1_.getFormat();
	let formatB1 = cellB1_.getFormat();
	let formatC1 = cellC1_.getFormat();

	formatA1.setColumnWidth(330);
	formatB1.setColumnWidth(230);
	formatC1.setColumnWidth(230);

	formatA1.getFill().setColor("#4472C4");

	formatA1.setRowHeight(80);

	cellB1_.setValue("· Protect the brand: eyes on the ground\n· Consistency of customer experience\n· Add value: not just a tool and critique\n· Speak their language as a business leader\n· Hold facts tightly and stories lightly");

	cellC1_.setValue("· Show gap: You won't change behavior unless you know what average is\n· Provide information that will blow away the directors\n· Ensure school is meeting goals and milestones" );

	formatB1.setWrapText(true);
	formatC1.setWrapText(true);


	// Get references to cells containing the important data
	let cellA1 = workbook.getWorksheet("Form Answer Report").getRange("C5");




	// Fetch the image from a URL.
	const link = "https://raw.githubusercontent.com/paulo-desouza/python-excel/main/celebree-logo.png";

	const response = await fetch(link);

	// Store the response as an ArrayBuffer, since it is a raw image file.
	const data = await response.arrayBuffer();

	// Convert the image data into a base64-encoded string.
	const image = convertToBase64(data);

	// Add the image to a worksheet.
	workbook.getWorksheet("Final Form").addImage(image);

	newSheet.getRange("A6:C6").merge();
	newSheet.getRange("A9:C9").merge();
	newSheet.getRange("A12:C12").merge();
	newSheet.getRange("A20:C20").merge();
	newSheet.getRange("A23:C23").merge();
	newSheet.getRange("A26:C26").merge();
	newSheet.getRange("A29:C29").merge();

	let cellA3_ = workbook.getWorksheet("Final Form").getRange("A3");

	let formatA3 = cellA3_.getFormat();

	formatA3.setRowHeight(40);

	cellA3_.setValue('SCHOOL VISIT SUMMARY');
	formatA3.getFont().setBold(true); // Change font to bold
	formatA3.getFont().setSize(24);
	formatA3.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
	formatA3.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

	let cellA4_ = workbook.getWorksheet("Final Form").getRange("A4");

	let formatA4 = cellA4_.getFormat();

	



	let cellB4_ = workbook.getWorksheet("Final Form").getRange("B4");

	let formatB4 = cellB4_.getFormat();




	let cellC4_ = workbook.getWorksheet("Final Form").getRange("C4");

	let formatC4 = cellC4_.getFormat();

	cellC4_.setValue('Visit Time:')


	let cellA5_ = workbook.getWorksheet("Final Form").getRange("A5");
	let formatA5 = cellA5_.getFormat();

	let cellB5_ = workbook.getWorksheet("Final Form").getRange("B5");
	let formatB5 = cellB5_.getFormat();

	let cellC5_ = workbook.getWorksheet("Final Form").getRange("C5");
	let formatC5 = cellC5_.getFormat();
	formatC5.getFill().setColor("#4472C4");
	formatB5.getFill().setColor("#4472C4");


	cellA5_.setValue('Purpose')
	formatA5.getFill().setColor("#4472C4");

	formatA5.getFont().setColor("#ffffff");
	formatA5.getFont().setBold(true); 

	let cellA8_ = workbook.getWorksheet("Final Form").getRange("A8");

	let formatA8 = cellA8_.getFormat();

	let cellB8_ = workbook.getWorksheet("Final Form").getRange("B8");
	let formatB8 = cellB8_.getFormat();

	let cellC8_ = workbook.getWorksheet("Final Form").getRange("C8");
	let formatC8 = cellC8_.getFormat();
	formatC8.getFill().setColor("#4472C4");
	formatB8.getFill().setColor("#4472C4");

	cellA8_.setValue('Objectives')
	formatA8.getFill().setColor("#4472C4");
	formatA8.getFont().setColor("#ffffff");
	formatA8.getFont().setBold(true); 

	let cellA11_ = workbook.getWorksheet("Final Form").getRange("A11");

	let formatA11 = cellA11_.getFormat();

	let cellB11_ = workbook.getWorksheet("Final Form").getRange("B11");
	let formatB11 = cellB11_.getFormat();

	let cellC11_ = workbook.getWorksheet("Final Form").getRange("C11");
	let formatC11 = cellC11_.getFormat();
	formatC11.getFill().setColor("#4472C4");
	formatB11.getFill().setColor("#4472C4");

	cellA11_.setValue('Brand Review')
	formatA11.getFill().setColor("#4472C4");
	formatA11.getFont().setColor("#ffffff");
	formatA11.getFont().setBold(true); 
	
	let cellA14_ = workbook.getWorksheet("Final Form").getRange("A14");

	let formatA14 = cellA14_.getFormat();
	let cellB14_ = workbook.getWorksheet("Final Form").getRange("B14");
	let formatB14 = cellB14_.getFormat();

	let cellC14_ = workbook.getWorksheet("Final Form").getRange("C14");
	let formatC14 = cellC14_.getFormat();
	formatC14.getFill().setColor("#4472C4");
	formatB14.getFill().setColor("#4472C4");


	cellA14_.setValue('Competency Ranking & Notes')
	formatA14.getFill().setColor("#4472C4");
	formatA14.getFont().setColor("#ffffff");
	formatA14.getFont().setBold(true); 

	let cellA19_ = workbook.getWorksheet("Final Form").getRange("A19");

	let formatA19 = cellA19_.getFormat();

	let cellB19_ = workbook.getWorksheet("Final Form").getRange("B19");
	let formatB19 = cellB19_.getFormat();

	let cellC19_ = workbook.getWorksheet("Final Form").getRange("C19");
	let formatC19 = cellC19_.getFormat();
	formatC19.getFill().setColor("#4472C4");
	formatB19.getFill().setColor("#4472C4");


	cellA19_.setValue('Goals & Milestones')
	formatA19.getFill().setColor("#4472C4");
	formatA19.getFont().setColor("#ffffff");
	formatA19.getFont().setBold(true); 

	let cellA22_ = workbook.getWorksheet("Final Form").getRange("A22");

	let formatA22 = cellA22_.getFormat();
	
	let cellB22_ = workbook.getWorksheet("Final Form").getRange("B22");
	let formatB22 = cellB22_.getFormat();

	let cellC22_ = workbook.getWorksheet("Final Form").getRange("C22");
	let formatC22 = cellC22_.getFormat();
	formatC22.getFill().setColor("#4472C4");
	formatB22.getFill().setColor("#4472C4");

	cellA22_.setValue('Notes from Visit Today')
	formatA22.getFill().setColor("#4472C4");
	formatA22.getFont().setColor("#ffffff");
	formatA22.getFont().setBold(true); 

	let cellA25_ = workbook.getWorksheet("Final Form").getRange("A25");

	let formatA25 = cellA25_.getFormat();

	let cellB25_ = workbook.getWorksheet("Final Form").getRange("B25");
	let formatB25 = cellB25_.getFormat();

	let cellC25_ = workbook.getWorksheet("Final Form").getRange("C25");
	let formatC25 = cellC25_.getFormat();
	formatC25.getFill().setColor("#4472C4");
	formatB25.getFill().setColor("#4472C4");



	cellA25_.setValue('Competency Supports Needed')
	formatA25.getFill().setColor("#4472C4");
	formatA25.getFont().setColor("#ffffff");
	formatA25.getFont().setBold(true); 



	let cellA28_ = workbook.getWorksheet("Final Form").getRange("A28");

	let formatA28 = cellA28_.getFormat();

	let cellB28_ = workbook.getWorksheet("Final Form").getRange("B28");
	let formatB28 = cellB28_.getFormat();

	let cellC28_ = workbook.getWorksheet("Final Form").getRange("C28");
	let formatC28 = cellC28_.getFormat();
	formatC28.getFill().setColor("#4472C4");
	formatB28.getFill().setColor("#4472C4");

	cellA28_.setValue('New Follow Up Actions and Due Dates to be Added to SAP')
	formatA28.getFill().setColor("#4472C4"); 
	formatA28.getFont().setColor("#ffffff");
	formatA28.getFont().setBold(true); 


	formatB1.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
	
	let edgeTop = formatB1.getRangeBorder(ExcelScript.BorderIndex.edgeTop);
	edgeTop.setStyle(ExcelScript.BorderLineStyle.dashDot);
	edgeTop.setWeight(ExcelScript.BorderWeight.thick);

	let edgeBottom = formatB1.getRangeBorder(ExcelScript.BorderIndex.edgeBottom);
	edgeBottom.setStyle(ExcelScript.BorderLineStyle.dashDot);
	edgeBottom.setWeight(ExcelScript.BorderWeight.thick);

	let edgeLeft = formatB1.getRangeBorder(ExcelScript.BorderIndex.edgeLeft);
	edgeLeft.setStyle(ExcelScript.BorderLineStyle.dashDot);
	edgeLeft.setWeight(ExcelScript.BorderWeight.thick);

	let edgeRight = formatC1.getRangeBorder(ExcelScript.BorderIndex.edgeRight);
	edgeRight.setStyle(ExcelScript.BorderLineStyle.dashDot);
	edgeRight.setWeight(ExcelScript.BorderWeight.thick);

	edgeTop = formatC1.getRangeBorder(ExcelScript.BorderIndex.edgeTop);
	edgeTop.setStyle(ExcelScript.BorderLineStyle.dashDot);
	edgeTop.setWeight(ExcelScript.BorderWeight.thick);

	edgeBottom = formatC1.getRangeBorder(ExcelScript.BorderIndex.edgeBottom);
	edgeBottom.setStyle(ExcelScript.BorderLineStyle.dashDot);
	edgeBottom.setWeight(ExcelScript.BorderWeight.thick);

	// set all cells row heights & wrap text & text alignment for receiving the data 
	// ROW NUMBERS: 6  9  12  15  16  17  20  23  26  29  

	let cellA6_ = workbook.getWorksheet("Final Form").getRange("A6");
	let formatA6 = cellA6_.getFormat();

	formatA6.setRowHeight(75);
	formatA6.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
	formatA6.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	formatA6.setWrapText(true);

	let cellA9_ = workbook.getWorksheet("Final Form").getRange("A9");
	let formatA9 = cellA9_.getFormat();

	formatA9.setRowHeight(75);
	formatA9.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
	formatA9.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	formatA9.setWrapText(true);

	let cellA12_ = workbook.getWorksheet("Final Form").getRange("A12");
	let formatA12 = cellA12_.getFormat();

	formatA12.setRowHeight(75);
	formatA12.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
	formatA12.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	formatA12.setWrapText(true);

	let cellA15_ = workbook.getWorksheet("Final Form").getRange("A15");
	let formatA15 = cellA15_.getFormat();

	formatA15.setRowHeight(75);
	formatA15.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
	formatA15.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	formatA15.setWrapText(true);

	let cellA16_ = workbook.getWorksheet("Final Form").getRange("A16");
	let formatA16 = cellA16_.getFormat();

	formatA16.setRowHeight(75);
	formatA16.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
	formatA16.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	formatA16.setWrapText(true);

	let cellA17_ = workbook.getWorksheet("Final Form").getRange("A17");
	let formatA17 = cellA17_.getFormat();

	formatA17.setRowHeight(75);
	formatA17.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
	formatA17.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	formatA17.setWrapText(true);

	
	let cellB15_ = workbook.getWorksheet("Final Form").getRange("B15");
	let formatB15 = cellB15_.getFormat();

	formatB15.setRowHeight(75);
	formatB15.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
	formatB15.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	formatB15.setWrapText(true);

	let cellB16_ = workbook.getWorksheet("Final Form").getRange("B16");
	let formatB16 = cellB16_.getFormat();

	formatB16.setRowHeight(75);
	formatB16.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
	formatB16.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	formatB16.setWrapText(true);

	let cellB17_ = workbook.getWorksheet("Final Form").getRange("B17");
	let formatB17 = cellB17_.getFormat();

	formatB17.setRowHeight(75);
	formatB17.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
	formatB17.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	formatB17.setWrapText(true);



	let cellC15_ = workbook.getWorksheet("Final Form").getRange("C15");
	let formatC15 = cellC15_.getFormat();

	formatC15.setRowHeight(75);
	formatC15.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
	formatC15.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	formatC15.setWrapText(true);

	let cellC16_ = workbook.getWorksheet("Final Form").getRange("C16");
	let formatC16 = cellC16_.getFormat();

	formatC16.setRowHeight(75);
	formatC16.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
	formatC16.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	formatC16.setWrapText(true);

	let cellC17_ = workbook.getWorksheet("Final Form").getRange("C17");
	let formatC17 = cellC17_.getFormat();

	formatC17.setRowHeight(75);
	formatC17.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
	formatC17.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	formatC17.setWrapText(true);



	let cellA20_ = workbook.getWorksheet("Final Form").getRange("A20");
	let formatA20 = cellA20_.getFormat();

	formatA20.setRowHeight(75);
	formatA20.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
	formatA20.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	formatA20.setWrapText(true);

	let cellA23_ = workbook.getWorksheet("Final Form").getRange("A23");
	let formatA23 = cellA23_.getFormat();

	formatA23.setRowHeight(75);
	formatA23.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
	formatA23.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	formatA23.setWrapText(true);

	let cellA26_ = workbook.getWorksheet("Final Form").getRange("A26");
	let formatA26 = cellA26_.getFormat();

	formatA26.setRowHeight(75);
	formatA26.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
	formatA26.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	formatA26.setWrapText(true);

	let cellA29_ = workbook.getWorksheet("Final Form").getRange("A29");
	let formatA29 = cellA29_.getFormat();

	formatA29.setRowHeight(75);
	
	formatA29.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

	formatA29.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
	formatA29.setWrapText(true);


	

	// get all the cells with data

	let franchise_id_cell = workbook.getWorksheet("Form Answer Report").getRange("F5");
	let consultant_cell = workbook.getWorksheet("Form Answer Report").getRange("G5");
	let purpose_cell = workbook.getWorksheet("Form Answer Report").getRange("I5");
	let objectives_cell = workbook.getWorksheet("Form Answer Report").getRange("K5");
	let br_cell = workbook.getWorksheet("Form Answer Report").getRange("M5");
	let rhs_cell = workbook.getWorksheet("Form Answer Report").getRange("O5");
	let abr_cell = workbook.getWorksheet("Form Answer Report").getRange("Q5");
	let pcte_cell = workbook.getWorksheet("Form Answer Report").getRange("S5");
	let cnd_cell = workbook.getWorksheet("Form Answer Report").getRange("U5");
	let cd_cell = workbook.getWorksheet("Form Answer Report").getRange("W5");
	let eie_cell = workbook.getWorksheet("Form Answer Report").getRange("Y5");
	let tbi_cell = workbook.getWorksheet("Form Answer Report").getRange("AA5");
	let pde_cell = workbook.getWorksheet("Form Answer Report").getRange("AC5");
	let pcte2_cell = workbook.getWorksheet("Form Answer Report").getRange("AE5");
	let goals_cell = workbook.getWorksheet("Form Answer Report").getRange("AG5");
	let notes_cell = workbook.getWorksheet("Form Answer Report").getRange("AI5");
	let comp_cell = workbook.getWorksheet("Form Answer Report").getRange("AK5");
	let followup_cell = workbook.getWorksheet("Form Answer Report").getRange("AM5");
	let date_cell = workbook.getWorksheet("Form Answer Report").getRange("C5");


	let cellC3_ = workbook.getWorksheet("Final Form").getRange("C3");
	cellC3_.setValue("Consultant: " + consultant_cell.getValue());


	cellB4_.setValue('Date: ' + date_cell.getValue());
	cellA4_.setValue('School: ' + franchise_id_cell.getValue());

	cellA6_.setValue(purpose_cell.getValue());
	cellA9_.setValue(objectives_cell.getValue());
	cellA12_.setValue(br_cell.getValue());

	cellA15_.setValue("RHS: " + rhs_cell.getValue());
	cellB15_.setValue("ABR: " + abr_cell.getValue());
	cellC15_.setValue("PCTE: " + pcte_cell.getValue());
	cellA16_.setValue("C&D: " + cnd_cell.getValue());
	cellB16_.setValue("CD: " + cd_cell.getValue());
	cellC16_.setValue("EIE: " + eie_cell.getValue());
	cellA17_.setValue("TBI: " + tbi_cell.getValue());
	cellB17_.setValue("PDE: " + pde_cell.getValue());
	cellC17_.setValue("PCTE: " + pcte2_cell.getValue());

	cellA20_.setValue(goals_cell.getValue());
	cellA23_.setValue(notes_cell.getValue());
	cellA26_.setValue(comp_cell.getValue());
	cellA29_.setValue(followup_cell.getValue());

	


	
}



/*
	// Get references to the cells they are going to
	let cellB2 = workbook.getWorksheet("Final Form").getRange("B4");


	// Move contents from A1 to B2
	cellB2.setValue(cellA1.getValue());


	// Change formatting of B2
	let formatB2 = cellB2.getFormat();
	formatB2.getFill().setColor("#4472C4"); // Change color
	formatB2.getFont().setBold(true); // Change font to bold
	formatB2.getFont().setSize(16);

	// Alignment and wrapping text

	formatB2.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
	formatB2.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	formatB2.setWrapText(true);

	let edgeTop = formatB2.getRangeBorder(ExcelScript.BorderIndex.edgeTop);
	edgeTop.setStyle(ExcelScript.BorderLineStyle.dashDot);
	edgeTop.setWeight(ExcelScript.BorderWeight.thick);

	let edgeBottom = formatB2.getRangeBorder(ExcelScript.BorderIndex.edgeBottom);
	edgeBottom.setStyle(ExcelScript.BorderLineStyle.dashDot);
	edgeBottom.setWeight(ExcelScript.BorderWeight.thick);

	let edgeLeft = formatB2.getRangeBorder(ExcelScript.BorderIndex.edgeLeft);
	edgeLeft.setStyle(ExcelScript.BorderLineStyle.dashDot);
	edgeLeft.setWeight(ExcelScript.BorderWeight.thick);

	let edgeRight = formatB2.getRangeBorder(ExcelScript.BorderIndex.edgeRight);
	edgeRight.setStyle(ExcelScript.BorderLineStyle.dashDot);
	edgeRight.setWeight(ExcelScript.BorderWeight.thick);

	formatB2.setColumnWidth(330);
	formatB2.setRowHeight(75);
}

*/


/**
 * Converts an ArrayBuffer containing a .png image into a base64-encoded string.
 */
function convertToBase64(input: ArrayBuffer) {
	const uInt8Array = new Uint8Array(input);
	const count = uInt8Array.length;

	// Allocate the necessary space up front.
	const charCodeArray = new Array(count) as string[];

	// Convert every entry in the array to a character.
	for (let i = count; i >= 0; i--) {
		charCodeArray[i] = String.fromCharCode(uInt8Array[i]);
	}

	// Convert the characters to base64.
	const base64 = btoa(charCodeArray.join(''));
	return base64;
}
