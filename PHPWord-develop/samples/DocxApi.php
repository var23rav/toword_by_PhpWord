<?php 
// define('CUSTOMIZE_FOR_RWR', true);
include_once 'Sample_Header.php';
// necessary code from Sample_Header.php
##########################################
// require_once __DIR__ . '/../src/PhpWord/Autoloader.php';

// date_default_timezone_set('UTC');

// /**
//  * Header file
//  */
// use PhpOffice\PhpWord\Autoloader;
// use PhpOffice\PhpWord\Settings;

// error_reporting(E_ALL);
// define('CLI', (PHP_SAPI == 'cli') ? true : false);
// define('EOL', CLI ? PHP_EOL : '<br />');
// define('SCRIPT_FILENAME', basename($_SERVER['SCRIPT_FILENAME'], '.php'));
// define('IS_INDEX', SCRIPT_FILENAME == 'index');

// Autoloader::register();
// Settings::loadConfig();

// // Set writers
// // $writers = array('Word2007' => 'docx', 'ODText' => 'odt', 'RTF' => 'rtf', 'HTML' => 'html', 'PDF' => 'pdf');
// $writers = array('Word2007' => 'docx'); // only for word document

// // Set PDF renderer
// if (null === Settings::getPdfRendererPath()) {
//     $writers['PDF'] = null;
// }

// // Return to the caller script when runs by CLI
// if (CLI) {
//     return;
// }

// // Set titles and names
// $pageHeading = str_replace('_', ' ', SCRIPT_FILENAME);
// $pageTitle = IS_INDEX ? 'Welcome to ' : "{$pageHeading} - ";
// $pageTitle .= 'PHPWord';
// $pageHeading = IS_INDEX ? '' : "<h1>{$pageHeading}</h1>";

// // Populate samples
// $files = '';
// if ($handle = opendir('.')) {
//     while (false !== ($file = readdir($handle))) {
//         if (preg_match('/^Sample_\d+_/', $file)) {
//             $name = str_replace('_', ' ', preg_replace('/(Sample_|\.php)/', '', $file));
//             $files .= "<li><a href='{$file}'>{$name}</a></li>";
//         }
//     }
//     closedir($handle);
// }



##############################################
// New Word Document
// echo date('H:i:s'), ' Create new PhpWord object', EOL;
$phpWord = new \PhpOffice\PhpWord\PhpWord();

// $PHPWord->addParagraphStyle('pJustify', array('align' => 'both', 'spaceBefore' => 0, 'spaceAfter' => 0, 'spacing' => 0));
// //add this style then append it to text below
// $section->addText('something', 'textstyle', 'pJustify');
// //the text behind this will be justified and will be in a new line, not in a new paragraph
// $section->addText('behind', 'textstyle', 'pJustify');

//----Style----
// $phpWord->addTitleStyle(1, array('size' => 16), array('numStyle' => 'hNum', 'numLevel' => 0));
// $phpWord->addTitleStyle(2, array('size' => 14), array('numStyle' => 'hNum', 'numLevel' => 1));
// $phpWord->addTitleStyle(3, array('size' => 12), array('numStyle' => 'hNum', 'numLevel' => 2));
//-------------
$phpWord->addTitleStyle(2, array('name'=>'Times New Roman', 'size'=>20, 'color'=>'000000','bold'=>true)); //h2

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
if( isset( $_POST['html_content'] ) ) {
	$html = $_POST['html_content'];
} else {
	$html = file_get_contents('mine.html');
}
// echo $_POST['html_content'];exit;
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

$docName = basename(__FILE__, '.php');

// Save file
echo write($phpWord, $docName, $writers);
// if( isset( $_POST['doc_name'] ) ) {
//     // echo saveFileByVar23($phpWord, $_POST['doc_name'], $writers);
//     echo write($phpWord, basename(__FILE__, '.php'), $writers);
//     // echo saveFileByVar23($phpWord, basename(__FILE__, '.php'), $writers);
// } else {
//     echo saveFileByVar23($phpWord, basename(__FILE__, '.php'), $writers);
// }

// $name = basename(__FILE__, '.php');
// $source = __DIR__ . "/resources/{$name}.docx";
// echo date('H:i:s'), " Reading contents from `{$source}`", EOL;
// $phpWord = \PhpOffice\PhpWord\IOFactory::load($source);
// // Save file
// echo write($phpWord, basename(__FILE__, '.php'), $writers);
if (!CLI) {
    include_once 'Sample_Footer.php';
}

$theExtractedFolder = lineBreakMSWordCompatibilityFix('results/' . $docName . '.docx' );
compressTheFolderToDocx('results/');

/*
1) avoid broken open and close tags
2) remove all <br> tag with &#xD;&#xA;  # carrier feed line feed
3) Remove the inner table for address slip
4) 

*/

// Same function as write() from sample_Header.php
function saveFileByVar23($phpWord, $filename, $writers)
{
    // $result = '';

    // Write documents
    foreach ($writers as $format => $extension) {
        // $result .= date('H:i:s') . " Write to {$format} format";
        if (null !== $extension) {
            $targetFile = __DIR__ . "/results/{$filename}.{$extension}";
            $phpWord->save($targetFile, $format);
        } else {
            // $result .= ' ... NOT DONE!';
        }
        // $result .= EOL;
    }

    // $result .= getEndingNotes($writers);

    // return $result;
    return $filename;
}


function lineBreakMSWordCompatibilityFix($docName) {
    $zip = new ZipArchive;
    if( $zip->open($docName) === TRUE ) {
        
        // Extracting the docx file to new folder with name of file
        $destFolder = rtrim($docName, '.docx');
        deleteFolder($destFolder);
        $zip->extractTo($destFolder);
        $zip->close();

        // Replacing all the #CRLF_BY_VAR23# which is added as the line break replacement with 
        // word xml line break <w:br/>
        $documentXmlFile = $destFolder . '/word/document.xml';
        if( file_exists($documentXmlFile) ) {
            $file_contents = file_get_contents($documentXmlFile);
            $file_contents = str_replace("#CRLF_BY_VAR23#","<w:br/>",$file_contents);
            file_put_contents($documentXmlFile,$file_contents);      
        }
        
        // Recompressing the folder into .docx file
        $filename = basename($docName, '.docx');
       // Get real path for our folder
        $rootPath = realpath('results/' . $filename);
        // Initialize archive object
        $zip = new ZipArchive();
        $zip->open($filename  . '123.docx', ZipArchive::CREATE | ZipArchive::OVERWRITE);

        // Create recursive directory iterator
        /** @var SplFileInfo[] $files */
        $files = new RecursiveIteratorIterator(
            new RecursiveDirectoryIterator($rootPath),
            RecursiveIteratorIterator::LEAVES_ONLY
        );

        foreach ($files as $name => $file)
        {
            // Skip directories (they would be added automatically)
            if (!$file->isDir())
            {
                // Get real and relative path for current file
                $filePath = $file->getRealPath();
                $relativePath = substr($filePath, strlen($rootPath) + 1);

                // Add current file to archive
                $zip->addFile($filePath, $relativePath);
            }
        }

        // Zip archive will be created only after closing object
        $zip->close();


    } else {
        echo  $docName . ' File extracition failed.';
    }
    return FALSE;
}

function deleteFolder($path)
{
    if (is_dir($path) === true)
    {
        $files = array_diff(scandir($path), array('.', '..'));

        foreach ($files as $file)
        {
            deleteFolder(realpath($path) . '/' . $file);
        }

        return rmdir($path);
    }

    else if (is_file($path) === true)
    {
        return unlink($path);
    }

    return false;
}

function compressTheFolderToDocx($folderPath) {
    $filname = rtrim($folderPath, '/') . 'docx';
    $zip =  new ZipArchive();
    $zip->open( $filname, ZipArchive::CREATE | ZipArchive::OVERWRITE);
}