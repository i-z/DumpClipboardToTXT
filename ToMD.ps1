$ext = [System.IO.Path]::GetExtension($args[0])
$newname = $args[0] -replace $ext, '.md'
'# ' + $newname + [Environment]::NewLine + ((gc $args[0]) -replace '\S+$','<img width="1000" src="$&">') | out-file $newname