param($path = $(throw "You must specify a config file"))
$global:appSettings = @{}
$config = [xml](get-content $path)
foreach ($addNode in $config.configuration.appsettings.add) {
 if ($addNode.Value.Contains(',')) {
  # Array case
  $value = $addNode.Value.Split(',')
  for ($i = 0; $i -lt $value.length; $i++) { 
    $value[$i] = $value[$i].Trim() 
  }
 }
 else {
  # Scalar case
  $value = $addNode.Value
 }
 $global:appSettings[$addNode.Key] = $value
}
