## Command Line arguments 

##updateconfig.ps1 FilePath XPath KeyName Value

## updateconfig.ps1 "C:\sample.config" "\\configuration\appSettings\add" "URL" "http://localhost/index.html"

$filename=$args[0]
$xpath=$args[1]
$keyname=$args[2]
$value=$args[3]


$xml= get-content $filename

$xml=[xml] $xml

$s=$xml.configuration.selectsinglenode("$xpath[@key='$keyname']")

#echo $s.getattribute("key")

$s.SetAttribute("value","$value")

$xml.save($filename)