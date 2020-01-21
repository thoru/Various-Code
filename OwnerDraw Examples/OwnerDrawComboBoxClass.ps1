# https://github.com/My-Random-Thoughts/Various-Code/blob/master/OwnerDraw%20Examples/OwnerDraw-ComboBox-Demo
# https://myrandomthoughts.co.uk/2017/08/ownerdraw-combobox/
# Tried to make the code as class. Supports multiple comboboxes 
# Class needs powershell 5

#Requires         -Version 4
#Set-StrictMode    -Version 2
#Remove-Variable * -ErrorAction SilentlyContinue
#Clear-Host

Add-Type -AssemblyName PresentationCore,PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Drawing.Font]$sysFont       = [System.Drawing.SystemFonts]::MessageBoxFont
[System.Windows.Forms.Application]::EnableVisualStyles()

class ComboBoxE {
    [System.Windows.Forms.ComboBox]$ComboBox
    [System.Collections.ArrayList]$comboItems
    [System.Windows.Forms.ImageList]$imgList
    [System.Windows.Forms.ImageList]$imgListWithIds

    ComboBoxE (){
        $this.imgList  = New-Object 'System.Windows.Forms.ImageList'
        $this.imgList.TransparentColor     = 'Transparent'

        $this.imgListWithIds  = New-Object 'System.Windows.Forms.ImageList'
        $this.imgListWithIds.TransparentColor     = 'Transparent'

        $this.ComboBox = New-Object System.Windows.Forms.ComboBox
        $this.comboItems = @{}
        
        $this| Add-Member -MemberType NoteProperty   -Name "ID" -Value ([guid]::NewGuid().ToString())
        $this.ComboBox | Add-Member -MemberType NoteProperty   -Name "ID" -Value ($this.ID)

        $this.ComboBox.DrawMode         = 'OwnerDrawFixed'
        $this.ComboBox.DropDownStyle    = 'DropDownList'
        $this.ComboBox.ItemHeight        = '20'
        $this.ComboBox.DropDownHeight   = (($this.ComboBox.ItemHeight * 10) + 2)

        $this.ComboBox.Add_DrawItem( {
                                         [System.Windows.Forms.DrawItemEventArgs]$e = $_

                                         # $this in this context is the combobox, use id stored in combobox to get back to the class
                                         $class      = $global:ComboBoxEA[$this.ID]
                                         $comboitems = $class.comboItems

                                         $e.DrawBackground()
                                            $e.DrawFocusRectangle()

                                            [System.Drawing.Rectangle]$bounds = $e.Bounds    
                                            If (($e.Index -gt -1) -and ($e.Index -lt $comboitems.Count)){
                                                $currItem = $comboitems[$e.Index]

                                                 if($currItem.Icon -match "^[\d\.]+$"){
                                                    $imgList  = $class.imgList
                                                    $currItem.Icon = [int] $currItem.Icon
                                                 }else{
                                                    $imgList  = $class.imgListWithIds  
                                                 }

                                                [int]                      $indent     = ($currItem.Indent * 14) 
                                                [System.Drawing.Image]     $icon       = $imgList.Images[$currItem.Icon]

                                                [System.Drawing.SolidBrush]$solidBrush = [System.Drawing.SolidBrush]$e.ForeColor

                                                If ($currItem.Enabled -eq $False) {
                                                   $icon = GreyScaleImage($icon)
                                                   $icon
                                                   $solidBrush.Color = [System.Drawing.SystemColors]::GrayText
                                                }

                                                $middle   = (($bounds.Top) + ((($bounds.Height) - ($icon.Height)) / 2))
                                                $iconRect = New-Object 'System.Drawing.RectangleF'((($bounds.Left) + 5 + $indent), $middle, $icon.Width, $icon.Width)
                                                $textRect = New-Object 'System.Drawing.RectangleF'((($bounds.Left) + ($iconRect.Width) + 9 + $indent), $bounds.Top, (($bounds.Width) - ($iconRect.Width) - 9 - $indent), $bounds.Height)
                                                $format   = New-Object 'System.Drawing.StringFormat'
                                                $format.Alignment     = [System.Drawing.StringAlignment]::Near
                                                $format.LineAlignment = [System.Drawing.StringAlignment]::Center

                                                If ($icon -ne $null) { $e.Graphics.DrawImage($icon, $iconRect) }
                                                $e.Graphics.DrawString($currItem.Text, $e.Font, $solidBrush, $textRect, $format)
                                            }


                                    } )
    }
    addComboItem([string] $Icon='0', [string] $Name="", [string] $Text="", [int] $Indent=0 , [boolean] $Enabled =$true ) { 
            # $icon can be number or chars. if a number DrawItem will lookup icons in imgList, if chars DrawItem will lookup in imgListWithIds
            $this.comboItems.Add( @{'Icon' = "$Icon"; 'Name' = $Name; 'Text' = $Text; 'Indent' = $Indent; 'Enabled' = $Enabled; } ) 
    }
    refreshCombobox(){
        $this.ComboBox.Items.Clear()
        $this.ComboBox.Items.AddRange($this.comboItems)
        $this.ComboBox.SelectedIndex = 0
    }
    addIMGtolist([string] $Path="") { 
            $tmp=[System.Drawing.Image]::FromFile($Path) 
            $this.imgList.Images.Add($tmp)#https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.imagelist.imagecollection?view=netframework-4.8
           
            $tmp#return
    }
    addIMGtolistWithID([string] $id="", [string] $Path="") { #https://powershell.org/forums/topic/class-methods-with-optional-parameter/
            $tmp=[System.Drawing.Image]::FromFile($Path) 
            $this.imgListWithIds.Images.Add($id, $tmp)#https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.imagelist.imagecollection?view=netframework-4.8
           
            $tmp#return
    }
    #Lookup a img in imagelist - $ComboBoxE2.imgList.Images[0]
    #Lookup a img in imagelist - $ComboBoxE2.imgListWithIds.Images["cut"]
}

Function GreyScaleImage ([System.Drawing.Image]$Image){# Used to "disable" an image (turn it into a greyscale image of itself)
    If ([string]::IsNullOrEmpty($Image) -eq $True) { Return $Image }
    [System.Drawing.Image]                  $newImage  = New-Object 'System.Drawing.Bitmap'($Image.Width, $Image.Height)
    [System.Drawing.Graphics]               $graphics  = [System.Drawing.Graphics]::FromImage($newImage)
    [System.Drawing.Imaging.ColorMatrix]    $matrix    = New-Object 'System.Drawing.Imaging.ColorMatrix'
    [System.Drawing.Imaging.ImageAttributes]$imgAttrib = New-Object 'System.Drawing.Imaging.ImageAttributes'
    $matrix.Matrix00 = '0.0'; $matrix.Matrix10 = '1.0'; $matrix.Matrix11 = '1.0'; $matrix.Matrix12 = '1.0'; $matrix.Matrix22 = '0.0'; $matrix.Matrix33 = '0.5'
    $imgAttrib.SetColorMatrix($matrix, [System.Drawing.Imaging.ColorMatrixFlag]::Default, [System.Drawing.Imaging.ColorAdjustType]::Bitmap)
    $graphics.DrawImage($Image, (New-Object 'System.Drawing.Rectangle'(0, 0, $Image.Width, $Image.Height)), 0, 0, $newImage.Width, $newImage.Height, [System.Drawing.GraphicsUnit]::Pixel, $imgAttrib)
    $graphics.Dispose()
    Return $newImage
}


#Using the combobox code
    # Use a array to manage all our imgcomboboxes
        $global:ComboBoxEA = @{}
        $ComboBoxE1 = [ComboBoxE]::New()
        $global:ComboBoxEA[$ComboBoxE1.ID]=$ComboBoxE1

        $ComboBoxE2 = [ComboBoxE]::New()
        $global:ComboBoxEA[$ComboBoxE2.ID]=$ComboBoxE2
 
    # Load icons
        $ComboBoxE1.addIMGtolist("D:\Powershell\appointment-new.png");
        $ComboBoxE1.addIMGtolist("D:\Powershell\go-top.png");
        $ComboBoxE1.addIMGtolist("D:\Powershell\edit-cut.png");
        $ComboBoxE1.addIMGtolist("D:\Powershell\format-indent-less.png");

        $ComboBoxE1.addIMGtolistWithID("appointment","D:\Powershell\appointment-new.png");
        $ComboBoxE1.addIMGtolistWithID("cut", "D:\Powershell\edit-cut.png");
        $ComboBoxE1.addIMGtolistWithID("go","D:\Powershell\go-top.png");

        $ComboBoxE2.imgList        = $ComboBoxE1.imgList        #Make both comboboxes use same imagelist
        $ComboBoxE2.imgListWithIds = $ComboBoxE1.imgListWithIds #Make both comboboxes use same imagelist

        #$ComboBoxE2.imgListWithIds.Images["cut"]

    # Adding items
        $ComboBoxE1.addComboItem( ($Icon='cut'), ($Name="Item1"), ($Text="Item1"), ($Indent=0), ($Enabled =$true) )
        $ComboBoxE1.addComboItem( ($Icon='go'),  ($Name="Item2"), ($Text="Item2"), ($Indent=0), ($Enabled =$false) )
        $ComboBoxE1.refreshCombobox()

        $ComboBoxE2.addComboItem( ($Icon='2'),   ($Name="Item3"), ($Text="Item3"), ($Indent=0), ($Enabled =$true) )
        $ComboBoxE2.addComboItem( ($Icon='3'),   ($Name="Item4"), ($Text="Item4"), ($Indent=1), ($Enabled =$true) )
        $ComboBoxE2.refreshCombobox()

    # Making form
        $MainFORM = New-Object 'System.Windows.Forms.Form'
        $MainFORM.FormBorderStyle       = 'FixedDialog'
        $MainFORM.MaximizeBox           = $False
        $MainFORM.MinimizeBox           = $False
        $MainFORM.ControlBox            = $True
        $MainFORM.Text                  = ' OwnerDraw cmoComboBox Control '
        $MainFORM.ShowInTaskbar         = $True
        $MainFORM.AutoScaleDimensions   = '6, 13'
        $MainFORM.AutoScaleMode         = 'None'
        $MainFORM.ClientSize            = '394, 147'
        $MainFORM.StartPosition         = 'CenterParent'

        $cmoComboBox1                   = $ComboBoxE1.ComboBox
        $cmoComboBox1.Location          = ' 12,  45'
        $cmoComboBox1.Width             = 370
        $cmoComboBox1.Font              = $sysFont
        $MainFORM.Controls.Add($cmoComboBox1)

        $cmoComboBox2                   = $ComboBoxE2.ComboBox
        $cmoComboBox2.Location          = ' 12,  75'
        $cmoComboBox2.Width             = 370
        $cmoComboBox2.Font              = $sysFont
        $MainFORM.Controls.Add($cmoComboBox2)


        $cmoComboBox1.Add_SelectedIndexChanged({
            If ($cmoComboBox1.SelectedIndex -lt 0) { Return }
            $selectedItem = $ComboBoxE1.comboItems[$cmoComboBox1.SelectedIndex]
            $MainFORM.Text = "Icon:$($selectedItem.Icon), Name:$($selectedItem.Name), Text:$($selectedItem.Text), Ident:$($selectedItem.Indent), Enabled:$($selectedItem.Enabled)"
        })
        $cmoComboBox2.Add_SelectedIndexChanged({
            If ($cmoComboBox2.SelectedIndex -lt 0) { Return }
            $selectedItem = $ComboBoxE2.comboItems[$cmoComboBox2.SelectedIndex]
            $MainFORM.Text = "Icon:$($selectedItem.Icon), Name:$($selectedItem.Name), Text:$($selectedItem.Text), Ident:$($selectedItem.Indent), Enabled:$($selectedItem.Enabled)"
        })

    ForEach ($control In $MainFORM.Controls) { $control.Font = $sysFont }
    $MainFORM.ShowDialog() | Out-Null
