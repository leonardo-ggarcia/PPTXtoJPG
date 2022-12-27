import { PowerShell } from "node-powershell";

/**
 * 
 * @param {*} inputFile  Ex: C:/Users/U01/
 * @param {*} outputFile Ex: C:/Users/U01/Pictures
 * @description Convert PPTX to JPG 
 */

function PPTXtoJPG(inputFile, outputFile){

    if(!inputFile && !outputFile){
        console.log("Params inputFile/outputFile are required")
        return;
    }

    try{
        PowerShell.$
        `           
            $app = New-Object -ComObj powerpoint.application
        
            $presentation = $app.Presentations.Open(${inputFile}, $true, $false, $false)
        
            $slideNumber=1
        
            foreach ($slide in $presentation.Slides){
        
            $slide = $presentation.Slides.Item($slideNumber)
            $slidePath = ${outputFile} +"\\" +$slideNumber +".jpg"
        
            $slideNumber += 1  
            $slide.Export($slidePath, "JPG")
            }
        
            $slide = $null
        
            $presentation.Close()
            $presentation = $null
        
            if($app.Windows.Count -eq 0)
            {
                $app.Quit()
            }
        
            $app = $null
            $slideNumber= $null
            $slidePath=$null
        `
    }
    catch(error){
        console.error("Unable to convert this file. ", error)
    }
  
}

export default PPTXtoJPG;