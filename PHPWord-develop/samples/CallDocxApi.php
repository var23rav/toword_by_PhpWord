<?php
function httpPost($url,$params) {

  $postData =  http_build_query($params);
  $ch = curl_init();  
  curl_setopt($ch,CURLOPT_URL,$url);
  curl_setopt($ch,CURLOPT_RETURNTRANSFER,true);
  curl_setopt($ch,CURLOPT_HEADER, false); 
  curl_setopt($ch, CURLOPT_POST, count($postData));
      curl_setopt($ch, CURLOPT_POSTFIELDS, $postData);    

  $output=curl_exec($ch);

  curl_close($ch);
  return $output;
 
}

// Work around phpword doesn't render html <br> tags
function sanitizeTheHtmlDataForDoc($htmlData) {
    $htmlData = str_ireplace([
            '<br />',
            '<br/>',
            '<br>',
        ], '<h7></h7>', $htmlData);//&#xA;&#xD;
    // $htmlData = br2nl($htmlData);
    return $htmlData;
}


$docName = isset($_GET['doc_name']) ? $_GET['doc_name'] : 'xda' ;

$htmlContent = file_get_contents('mine.html');
$sanitizedHtmlContent = sanitizeTheHtmlDataForDoc($htmlContent);
$params = array(
   'html_content' => $sanitizedHtmlContent,
   'doc_name' => $docName,
);
 
echo httpPost("http://localhost/PHPWord-develop/samples/DocxApi.php",$params);