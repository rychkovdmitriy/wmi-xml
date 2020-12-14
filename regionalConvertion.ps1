     $culture=Get-Culture
     $culture.NumberFormat.NumberDecimalSeparator

     $value = "50,3"
     $styles = [System.Globalization.NumberStyles]::Float
     $provider = [System.Globalization.CultureInfo]::CreateSpecificCulture("ru-RU")
     $double = [double]::Parse($value, $styles, $provider)
     $double


     $value = "50.4"
     $styles = [System.Globalization.NumberStyles]::Float
     $provider = [System.Globalization.CultureInfo]::CreateSpecificCulture("en-US")
     $double = [double]::Parse($value, $styles, $provider)
     $double


     if($culture.NumberFormat.NumberDecimalSeparator -eq ",")
     {
        $styles = [System.Globalization.NumberStyles]::Float
        $provider = [System.Globalization.CultureInfo]::CreateSpecificCulture("en-US")
        $double = [double]::Parse($value, $styles, $provider)
        $double
     }
