Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Test-ValidFileName
{
    param([string]$FileName)

    $IndexOfInvalidChar = $FileName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars())

    # IndexOfAny() returns the value -1 to indicate no such character was found
    return $IndexOfInvalidChar -eq -1
}

$htmlBegin = @"
<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8">
  </head>
  <body>
"@

$htmlEnd = @"
</body>
</html>
"@

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Data Entry Form'
$form.Size = New-Object System.Drawing.Size(300, 250)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(55,160)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)


$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(170,160)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$decorateWithImgCheck = New-Object System.Windows.Forms.CheckBox
$decorateWithImgCheck.Location = New-Object System.Drawing.Point(10,120)
$decorateWithImgCheck.Size = New-Object System.Drawing.Size(130,40)
$decorateWithImgCheck.Text = 'Decorate with img'
$form.Controls.Add($decorateWithImgCheck)

$htmlOutput = New-Object System.Windows.Forms.CheckBox
$htmlOutput.Location = New-Object System.Drawing.Point(150,120)
$htmlOutput.Size = New-Object System.Drawing.Size(130,40)
$htmlOutput.Text = 'HTML'
$form.Controls.Add($htmlOutput)

$content = Get-Clipboard
$clipboardText = [System.Windows.Forms.Clipboard]::GetText()
$preview = ""
if ($clipboardText.Length -gt 150) {
    $preview = $clipboardText.Substring(0, 150)
} else {
    $preview = $clipboardText
}

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,60)
#$label.Text = [string]::Format("Cipboard size: {0}. Begin with:'{1}'", $clipboardText.Length, $preview)
$label.Text = [string]::Format("Cipboard size: {0:N0} {1} {2}", $clipboardText.Length, [Environment]::NewLine, $preview)
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,100)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox)

$form.Topmost = $true

$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $x = $textBox.Text
    if(-not (Test-ValidFileName $x))
    {
        [System.IO.Path]::GetInvalidFileNameChars() | Foreach-Object {
            $x = $x.Replace($_, ' ');
        }
    }
    $ext = '.txt'
    if ($decorateWithImgCheck.Checked -eq $true) {
        $out = New-Object System.Text.StringBuilder
        $imageTagBegin = '![]('
        $imageTagEnd = ')' + [Environment]::NewLine + [Environment]::NewLine
        if ($htmlOutput.Checked -eq $true) {
            $out.Append($htmlBegin)
            $imageTagBegin = '<p><img src="'
            $imageTagEnd = '"></p>' + [Environment]::NewLine
        }
        ForEach ($line in $content){
            $out.Append($imageTagBegin)
            $out.Append($line)
            $out.Append($imageTagEnd)
        }
        if ($htmlOutput.Checked -eq $true) {
            $out.Append($htmEnd)
            $out.ToString() | out-file -LiteralPath ($x + ".html") -encoding utf8
        } else {
            $out.ToString() | out-file -LiteralPath ($x + ".md") -encoding utf8
        }
        
    } else {
        $content | out-file -LiteralPath ($x + ".txt") -encoding utf8
    }
}