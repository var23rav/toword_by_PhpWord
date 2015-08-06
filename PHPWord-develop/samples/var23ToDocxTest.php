<?php
include_once 'Sample_Header.php';

// New Word Document
// echo date('H:i:s'), ' Create new PhpWord object', EOL;
$phpWord = new \PhpOffice\PhpWord\PhpWord();

$section = $phpWord->addSection();

//------------------Header n Footer-----------------------
// Add header for all other pages
$subsequent = $section->addHeader();
// $subsequent->addText(htmlspecialchars('Subsequent pages in Section 1 will Have this!', ENT_COMPAT, 'UTF-8'));

// Add footer
$footer = $section->addFooter();
$footer->addPreserveText(htmlspecialchars('Page {PAGE} of {NUMPAGES}.', ENT_COMPAT, 'UTF-8'), null, array('alignment' => 'center'));
//-----------------------------------------





// $normalFont['name'] = 'Times New Roman';
// $normalFont['size'] = 12;

// $normalUnderlineFont = $normalFont;
// $normalUnderlineFont['size'] = 10;

// $fontStyle['name'] = 'Times New Roman';
// $fontStyle['size'] = 20;
// $fontStyle['bold'] = true;

// $textrun = $section->addTextRun();
// $textrun->addText(htmlspecialchars('Dit is het scheidingsplan van:', ENT_COMPAT, 'UTF-8'), $fontStyle);

// $heading_n = 'Communiceren';
// $heading_u = 'Evaluatie ';
// $content = 'We vinden het belangrijk dat we tot een goede oplossing komen over waar ieder van ons gaat wonen. Daarom spreken we af dat (Fill in: the person`s name that remains in the rental property) continues to reside in the home. (Fill in the name of the person who stays in the family house) is going to pay the rent from (fill in: date) .We ask the landlord as of the date the lease on (fill in the name of the person in the rental remains) programs to put name.';



// $section->addTextBreak(2);

// for($i=0;$i<15;$i++) {
// 	$textrun = $section->addTextRun();
// 	$textrun->addText(htmlspecialchars($heading_n, ENT_COMPAT, 'UTF-8'), $normalFont);

// 	$textrun = $section->addTextRun();
// 	$textrun->addText(htmlspecialchars($heading_u, ENT_COMPAT, 'UTF-8'), $normalUnderlineFont);

// 	$textrun = $section->addTextRun();
// 	$textrun->addText(htmlspecialchars($content, ENT_COMPAT, 'UTF-8'));
// 	$section->addTextBreak(2);
// }



$html = file_get_contents('mine.html');
\PhpOffice\PhpWord\Shared\Html::addHtml($section, $html);





//*******************Address slip********************
//Dummy table without data to avoid problem with all other table
$table = $section->addTable();
$table->addRow();
$table->addCell(1750)->addText(htmlspecialchars("", ENT_COMPAT, 'UTF-8'));

$addressSlip = [
	'te ..............................................',
	'op datum ........................................',
	'.................................................',
	'(Naam) ..........................................',
	'(Handtekening) ..................................'
];

$section->addTextBreak(1);

$table = $section->addTable();
for($r = 0; $r < count($addressSlip); $r++) { // Loop through rows
	// Add row
	$table->addRow(900);
	for ($c = 0; $c < 2; $c++) {
		$table->addCell(4500)->addText(htmlspecialchars($addressSlip[$r], ENT_COMPAT, 'UTF-8'));
	}
}
//***********************************************



// Save file
echo write($phpWord, basename(__FILE__, '.php'), $writers);
if (!CLI) {
    include_once 'Sample_Footer.php';
}



/*
1) avoid broken open and close tags
2) remove all <br> tag with &#xD;&#xA;  # carrier feed line feed
3) Remove the inner table for address slip
4) 

*/

