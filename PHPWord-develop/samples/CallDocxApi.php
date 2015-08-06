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


$docName = isset($_GET['doc_name']) ? $_GET['doc_name'] : 'xda' ;

$htmlContent = file_get_contents('mine.html');

$params = array(
   'html_content' => $htmlContent,
   'doc_name' => $docName,
);
echo httpPost("http://localhost/toword_by_PhpWord/PHPWord-develop/samples/DocxApi.php",$params);