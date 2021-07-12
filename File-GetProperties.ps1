<#

 Microsoft provides programming examples for illustration only, without warranty either expressed or
 implied, including, but not limited to, the implied warranties of merchantability and/or fitness 
 for a particular purpose. 
 
 This sample assumes that you are familiar with the programming language being demonstrated and the 
 tools used to create and debug procedures. Microsoft support professionals can help explain the 
 functionality of a particular procedure, but they will not modify these examples to provide added 
 functionality or construct procedures to meet your specific needs. if you have limited programming 
 experience, you may want to contact a Microsoft Certified Partner or the Microsoft fee-based consulting 
 line at (800) 936-5200. 

 
 ----------------------------------------------------------
 Purpose:
 ----------------------------------------------------------

 The purpose of this script is to itterate through all the MP4 files in the specified directory and extract the detailed attributes of the file.

 The results will be saved to the log directory file specified.

==============================================================#>
$LogDirectory = "C:\temp\FilePropertiesOutput\output.csv"

function AddFileDetails
{
    Param
    (
        [String] [Parameter(Mandatory=$true, ValueFromPipeline=$true)] $KeyName,
        [String] [Parameter(Mandatory=$true, ValueFromPipeline=$true)] $KeyValue
    )
    $FileDetailsItem = New-Object PSObject
    $FileDetailsItem | Add-Member -MemberType NoteProperty -Name $KeyName -Value $KeyValue

    $FileDetails += $FileDetailsItem

    $FileDetails=@(
        [pscustomobject]@{
            KeyName=$KeyName 
            KeyValue=$KeyValue})

    $FileDetails | export-csv -Path $LogDirectory -NoTypeInformation -Append
}

Function Get-MP4MetaData
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([Psobject])]
    Param
    (
        [String] [Parameter(Mandatory=$true, ValueFromPipeline=$true)] $Directory
    )

   Begin
   {
        $shell = New-Object -ComObject "Shell.Application"
   }
   Process
  {
        Foreach($Dir in $Directory)
        {
            $ObjDir = $shell.NameSpace($Dir)
            
            $Files = Get-ChildItem $Dir -Filter '*.mp4'

            Foreach($File in $Files)
            {
                $ObjFile = $ObjDir.parsename($File.Name)
                $MetaData = @{}

                for ($i=1; $i -le 1000; $i++) 
                {
                    If($ObjDir.GetDetailsOf($ObjFile, $i)) #To avoid empty values
                    {
                        $MetaData[$($i.ToString())] = $ObjDir.GetDetailsOf($ObjFile, $i)
                    }

                }

                foreach ($item in $MetaData.GetEnumerator() | Sort-Object Name) {


                    switch ($item.Key) {

                        1 { AddFileDetails -KeyName "Size" -KeyValue $item.Value }
                        10 {AddFileDetails -KeyName "Owner" -KeyValue $item.Value}
                        11 {AddFileDetails -KeyName "FileType" -KeyValue $item.Value}
                        13 {AddFileDetails -KeyName "ContributingArtist" -KeyValue $item.Value}
                        15 {AddFileDetails -KeyName "Year" -KeyValue $item.Value}
                        16 {AddFileDetails -KeyName "Genre" -KeyValue $item.Value}
                        164 {AddFileDetails -KeyName "FileExtension" -KeyValue $item.Value}
                        165 {AddFileDetails -KeyName "Name" -KeyValue $item.Value}
                        169 {AddFileDetails -KeyName "UNKValue1" -KeyValue $item.Value}
                        17 {AddFileDetails -KeyName "Conductors" -KeyValue $item.Value}
                        18 {AddFileDetails -KeyName "Tags" -KeyValue $item.Value}
                        187 {AddFileDetails -KeyName "Protected" -KeyValue $item.Value}
                        19 {AddFileDetails -KeyName "Rating" -KeyValue $item.Value}
                        190 {AddFileDetails -KeyName "Folder" -KeyValue $item.Value}
                        191 {AddFileDetails -KeyName "FolderPath" -KeyValue $item.Value}
                        192 {AddFileDetails -KeyName "ParentFolder" -KeyValue $item.Value}
                        194 {AddFileDetails -KeyName "FullPath" -KeyValue $item.Value}
                        196 {AddFileDetails -KeyName "ItemType" -KeyValue $item.Value}
                        2 {AddFileDetails -KeyName "ItemType2" -KeyValue $item.Value}
                        20 {AddFileDetails -KeyName "ContributingArtist2" -KeyValue $item.Value}
                        202 {AddFileDetails -KeyName "UNKValue3" -KeyValue $item.Value}
                        208 {AddFileDetails -KeyName "MediaCreated" -KeyValue $item.Value}
                        21 {AddFileDetails -KeyName "Title1" -KeyValue $item.Value}
                        28 {AddFileDetails -KeyName "Title" -KeyValue $item.Value}
                        210 {AddFileDetails -KeyName "EncodedBy" -KeyValue $item.Value}
                        212 {AddFileDetails -KeyName "Producers" -KeyValue $item.Value}
                        213 {AddFileDetails -KeyName "Publisher" -KeyValue $item.Value}
                        215 {AddFileDetails -KeyName "Subtitle" -KeyValue $item.Value}
                        217 {AddFileDetails -KeyName "Writers" -KeyValue $item.Value}
                        24 {AddFileDetails -KeyName "Comments" -KeyValue $item.Value}
                        242 {AddFileDetails -KeyName "BeatsPerMinute" -KeyValue $item.Value}
                        243 {AddFileDetails -KeyName "Composers" -KeyValue $item.Value}
                        246 {AddFileDetails -KeyName "InitialKey" -KeyValue $item.Value}
                        248 {AddFileDetails -KeyName "Mood" -KeyValue $item.Value}
                        250 {AddFileDetails -KeyName "Period" -KeyValue $item.Value}
                        252 {AddFileDetails -KeyName "ParentalRating" -KeyValue $item.Value}
                        254 {AddFileDetails -KeyName "UNKValue4" -KeyValue $item.Value}
                        27 {AddFileDetails -KeyName "Length" -KeyValue $item.Value}
                        279 {AddFileDetails -KeyName "Subtitle2" -KeyValue $item.Value}
                        28 {AddFileDetails -KeyName "BitRate" -KeyValue $item.Value}
                        29 {AddFileDetails -KeyName "UNKValue4" -KeyValue $item.Value}
                        295 {AddFileDetails -KeyName "Sharing" -KeyValue $item.Value}
                        296 {AddFileDetails -KeyName "UNKValue5" -KeyValue $item.Value}
                        3 {AddFileDetails -KeyName "DateModified" -KeyValue $item.Value}
                        311 {AddFileDetails -KeyName "UNKValue6" -KeyValue $item.Value}
                        312 {AddFileDetails -KeyName "Directors" -KeyValue $item.Value}
                        313 {AddFileDetails -KeyName "DataRate" -KeyValue $item.Value}
                        314 {AddFileDetails -KeyName "FrameHeight" -KeyValue $item.Value}
                        315 {AddFileDetails -KeyName "FrameRate" -KeyValue $item.Value}
                        316 {AddFileDetails -KeyName "FrameWidth" -KeyValue $item.Value}
                        317 {AddFileDetails -KeyName "UNKValue7" -KeyValue $item.Value}
                        318 {AddFileDetails -KeyName "UNKValue8" -KeyValue $item.Value}
                        319 {AddFileDetails -KeyName "UNKValue9" -KeyValue $item.Value}
                        320 {AddFileDetails -KeyName "UNKValue10" -KeyValue $item.Value}
                        4 {AddFileDetails -KeyName "CreatedDate" -KeyValue $item.Value}
                        5 {AddFileDetails -KeyName "UNKValue11" -KeyValue $item.Value}
                        57 {AddFileDetails -KeyName "UNKValue12" -KeyValue $item.Value}
                        6 {AddFileDetails -KeyName "Attributes" -KeyValue $item.Value}
                        61 {AddFileDetails -KeyName "Computer" -KeyValue $item.Value}
                        9 {AddFileDetails -KeyName "UNKValue13" -KeyValue $item.Value}

                    }
            }
        }
    }
  }
}


Get-MP4MetaData -Directory "C:\temp\FileProperties"
