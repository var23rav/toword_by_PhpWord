<?php
include_once 'Sample_Header.php';

// Read contents
$name = basename(__FILE__, '.php');
$source = __DIR__ . "/resources/{$name}.docx";

echo date('H:i:s'), " Reading contents from `{$source}`", EOL;
$phpWord = \PhpOffice\PhpWord\IOFactory::load($source);


$source = __DIR__ . "/resources/Text_var23.docx";
$phpWord = \PhpOffice\PhpWord\IOFactory::load($source);
$sections = $phpWord->getSections();
$section = $sections[0]; // le document ne contient qu'une section
// $arrays = $section->getElements();
// arr($sections);
foreach($sections as $section){
	$footers = $section->getFooters();
	echo '</p>-----------------------------------------------';
	foreach ($footers as $footer) {
		$elements = $footer->getElements();
		$element1 = $elements[0]; var_dump($element1->getText());
		// $element1->setText("This is my text addition - old part: " . $element1->getText());
	// // note that the first index is 1 here (not 0)
	// for($i=1; $i < count($footers); $i++){
	// 	$footer = $footers[$i];
		// $header1 = $headers[1]; // note that the first index is 1 here (not 0)

		// $elements = $header1->getElements();
		// $element1 = $elements[0]; // and first index is 0 here normally

		// // for example manipulating simple text information ($element1 is instance of Text object)
		// $element1->setText("This is my text addition - old part: " . $element1->getText());
	}
}

// arr($arrays);

exit;
// Save file
echo write($phpWord, basename(__FILE__, '.php'), $writers);
if (!CLI) {
    include_once 'Sample_Footer.php';
}

function arr($array){
	echo '<pre>';
	print_r($array);
	echo '</pre>';

}