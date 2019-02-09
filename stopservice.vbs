'Tanium File Version:7.1.7.0010
'@INCLUDE=i18n/UTF8Decode.vbs

Option Explicit

Dim colArguments,strService,objShell,intReturn,strCommand

Set colArguments = WScript.Arguments

If WScript.Arguments.Count < 1 Then
	WScript.Echo "Usage - " & wscript.ScriptName & " <service to stop>"
	WScript.Quit
End If


strService = Trim(UTF8Decode(colArguments.Item(0)))

If IsInvalidFileName(strService) Then
	WScript.Echo "Invalid input: " & strService
	WScript.Quit
End If


Set objShell = CreateObject("WScript.Shell")

strCommand = "net stop " &Chr(34)&strService&Chr(34)
WScript.Echo "About to run: " & strCommand
intReturn = objShell.Run(strCommand,0,True)

If intReturn = 0 Then
	WScript.Echo "Success"
Else
	WScript.Echo "Failed"
End If

Function IsInvalidFileName(strFile)
' returns True if the string could be a valid windows file name
' in that it doesn't contain \ / : * ? " < > |

	Dim bIsNotValidFile : bIsNotValidFile = False
	Dim arrInvalidChars,strChar
	
	arrInvalidChars = Array("\", "/", ":", "*", "?", Chr(34), "<", ">", "|")
	For Each strChar In arrInvalidChars
		If InStr(strFile,strChar) > 0 Then ' bad character
			bIsNotValidFile = True
		End If
	Next
	
	IsInvalidFileName = bIsNotValidFile
End Function 'IsInvalidFileName
	
'------------ INCLUDES after this line. Do not edit past this point -----
'- Begin file: i18n/UTF8Decode.vbs
'========================================
' UTF8Decode
'========================================
' Used to convert the UTF-8 style parameters passed from 
' the server to sensors in sensor parameters.
' This function should be used to safely pass non english input to sensors.

' To include this file, copy/paste: INCLUDE=i18n/UTF8Decode.vbs


Function UTF8Decode(str)
    Dim arraylist(), strLen, i, sT, val, depth, sR
    Dim arraysize
    arraysize = 0
    strLen = Len(str)
    for i = 1 to strLen
        sT = mid(str, i, 1)
        if sT = "%" then
            if i + 2 <= strLen then
                Redim Preserve arraylist(arraysize + 1)
                arraylist(arraysize) = cbyte("&H" & mid(str, i + 1, 2))
                arraysize = arraysize + 1
                i = i + 2
            end if
        else
            Redim Preserve arraylist(arraysize + 1)
            arraylist(arraysize) = asc(sT)
            arraysize = arraysize + 1
        end if
    next
    depth = 0
    for i = 0 to arraysize - 1
		Dim mybyte
        mybyte = arraylist(i)
        if mybyte and &h80 then
            if (mybyte and &h40) = 0 then
                if depth = 0 then
                    Err.Raise 5
                end if
                val = val * 2 ^ 6 + (mybyte and &h3f)
                depth = depth - 1
                if depth = 0 then
                    sR = sR & chrw(val)
                    val = 0
                end if
            elseif (mybyte and &h20) = 0 then
                if depth > 0 then Err.Raise 5
                val = mybyte and &h1f
                depth = 1
            elseif (mybyte and &h10) = 0 then
                if depth > 0 then Err.Raise 5
                val = mybyte and &h0f
                depth = 2
            else
                Err.Raise 5
            end if
        else
            if depth > 0 then Err.Raise 5
            sR = sR & chrw(mybyte)
        end if
    next
    if depth > 0 then Err.Raise 5
    UTF8Decode = sR
End Function
'- End file: i18n/UTF8Decode.vbs

'' SIG '' Begin signature block
'' SIG '' MIIX9AYJKoZIhvcNAQcCoIIX5TCCF+ECAQExCzAJBgUr
'' SIG '' DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
'' SIG '' gjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIBAAIB
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFHlb0HuKG8SW
'' SIG '' Ys1wh9WwT7UQQGIioIITKTCCA+4wggNXoAMCAQICEH6T
'' SIG '' 6/t8xk5Z6kuad9QG/DswDQYJKoZIhvcNAQEFBQAwgYsx
'' SIG '' CzAJBgNVBAYTAlpBMRUwEwYDVQQIEwxXZXN0ZXJuIENh
'' SIG '' cGUxFDASBgNVBAcTC0R1cmJhbnZpbGxlMQ8wDQYDVQQK
'' SIG '' EwZUaGF3dGUxHTAbBgNVBAsTFFRoYXd0ZSBDZXJ0aWZp
'' SIG '' Y2F0aW9uMR8wHQYDVQQDExZUaGF3dGUgVGltZXN0YW1w
'' SIG '' aW5nIENBMB4XDTEyMTIyMTAwMDAwMFoXDTIwMTIzMDIz
'' SIG '' NTk1OVowXjELMAkGA1UEBhMCVVMxHTAbBgNVBAoTFFN5
'' SIG '' bWFudGVjIENvcnBvcmF0aW9uMTAwLgYDVQQDEydTeW1h
'' SIG '' bnRlYyBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIENBIC0g
'' SIG '' RzIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
'' SIG '' AQCxrLNJVEuXHBIK2CV5kSJXKm/cuCbEQ3Nrwr8uUFr7
'' SIG '' FMJ2jkMBJUO0oeJF9Oi3e8N0zCLXtJQAAvdN7b+0t0Qk
'' SIG '' a81fRTvRRM5DEnMXgotptCvLmR6schsmTXEfsTHd+1Fh
'' SIG '' AlOmqvVJLAV4RaUvic7nmef+jOJXPz3GktxK+Hsz5HkK
'' SIG '' +/B1iEGc/8UDUZmq12yfk2mHZSmDhcJgFMTIyTsU2sCB
'' SIG '' 8B8NdN6SIqvK9/t0fCfm90obf6fDni2uiuqm5qonFn1h
'' SIG '' 95hxEbziUKFL5V365Q6nLJ+qZSDT2JboyHylTkhE/xni
'' SIG '' RAeSC9dohIBdanhkRc1gRn5UwRN8xXnxycFxAgMBAAGj
'' SIG '' gfowgfcwHQYDVR0OBBYEFF+a9W5czMx0mtTdfe8/2+xM
'' SIG '' gC7dMDIGCCsGAQUFBwEBBCYwJDAiBggrBgEFBQcwAYYW
'' SIG '' aHR0cDovL29jc3AudGhhd3RlLmNvbTASBgNVHRMBAf8E
'' SIG '' CDAGAQH/AgEAMD8GA1UdHwQ4MDYwNKAyoDCGLmh0dHA6
'' SIG '' Ly9jcmwudGhhd3RlLmNvbS9UaGF3dGVUaW1lc3RhbXBp
'' SIG '' bmdDQS5jcmwwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDgYD
'' SIG '' VR0PAQH/BAQDAgEGMCgGA1UdEQQhMB+kHTAbMRkwFwYD
'' SIG '' VQQDExBUaW1lU3RhbXAtMjA0OC0xMA0GCSqGSIb3DQEB
'' SIG '' BQUAA4GBAAMJm495739ZMKrvaLX64wkdu0+CBl03X6ZS
'' SIG '' nxaN6hySCURu9W3rWHww6PlpjSNzCxJvR6muORH4KrGb
'' SIG '' sBrDjutZlgCtzgxNstAxpghcKnr84nodV0yoZRjpeUBi
'' SIG '' JZZux8c3aoMhCI5B6t3ZVz8dd0mHKhYGXqY4aiISo1EZ
'' SIG '' g362MIIEozCCA4ugAwIBAgIQDs/0OMj+vzVuBNhqmBsa
'' SIG '' UDANBgkqhkiG9w0BAQUFADBeMQswCQYDVQQGEwJVUzEd
'' SIG '' MBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAu
'' SIG '' BgNVBAMTJ1N5bWFudGVjIFRpbWUgU3RhbXBpbmcgU2Vy
'' SIG '' dmljZXMgQ0EgLSBHMjAeFw0xMjEwMTgwMDAwMDBaFw0y
'' SIG '' MDEyMjkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMR0wGwYD
'' SIG '' VQQKExRTeW1hbnRlYyBDb3Jwb3JhdGlvbjE0MDIGA1UE
'' SIG '' AxMrU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2aWNl
'' SIG '' cyBTaWduZXIgLSBHNDCCASIwDQYJKoZIhvcNAQEBBQAD
'' SIG '' ggEPADCCAQoCggEBAKJjCzlEuLsjp0RJuw7/ofBhClOT
'' SIG '' sJjbrSwPSsVu/4Y8U1UPFc4EPyv9qZaW2b5heQtbyUyG
'' SIG '' duXgQ0sile7CK0PBn9hotI5AT+6FOLkRxSPyZFjwFTJv
'' SIG '' TlehroikAtcqHs1L4d1j1ReJMluwXplaqJ0oUA4X7pbb
'' SIG '' YTtFUR3PElYLkkf8q672Zj1HrHBy55LnX80QucSDZJQZ
'' SIG '' vSWA4ejSIqXQugJ6oXeTW2XD7hd0vEGGKtwITIySjJEt
'' SIG '' nndEH2jWqHR32w5bMotWizO92WPISZ06xcXqMwvS8aMb
'' SIG '' 9Iu+2bNXizveBKd6IrIkri7HcMW+ToMmCPsLvalPmQjh
'' SIG '' EChyqs0CAwEAAaOCAVcwggFTMAwGA1UdEwEB/wQCMAAw
'' SIG '' FgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/
'' SIG '' BAQDAgeAMHMGCCsGAQUFBwEBBGcwZTAqBggrBgEFBQcw
'' SIG '' AYYeaHR0cDovL3RzLW9jc3Aud3Muc3ltYW50ZWMuY29t
'' SIG '' MDcGCCsGAQUFBzAChitodHRwOi8vdHMtYWlhLndzLnN5
'' SIG '' bWFudGVjLmNvbS90c3MtY2EtZzIuY2VyMDwGA1UdHwQ1
'' SIG '' MDMwMaAvoC2GK2h0dHA6Ly90cy1jcmwud3Muc3ltYW50
'' SIG '' ZWMuY29tL3Rzcy1jYS1nMi5jcmwwKAYDVR0RBCEwH6Qd
'' SIG '' MBsxGTAXBgNVBAMTEFRpbWVTdGFtcC0yMDQ4LTIwHQYD
'' SIG '' VR0OBBYEFEbGaaMOShQe1UzaUmMXP142vA3mMB8GA1Ud
'' SIG '' IwQYMBaAFF+a9W5czMx0mtTdfe8/2+xMgC7dMA0GCSqG
'' SIG '' SIb3DQEBBQUAA4IBAQB4O7SRKgBM8I9iMDd4o4QnB28Y
'' SIG '' st4l3KDUlAOqhk4ln5pAAxzdzuN5yyFoBtq2MrRtv/Qs
'' SIG '' JmMz5ElkbQ3mw2cO9wWkNWx8iRbG6bLfsundIMZxD82V
'' SIG '' dNy2XN69Nx9DeOZ4tc0oBCCjqvFLxIgpkQ6A0RH83Vx2
'' SIG '' bk9eDkVGQW4NsOo4mrE62glxEPwcebSAe6xp9P2ctgwW
'' SIG '' K/F/Wwk9m1viFsoTgW0ALjgNqCmPLOGy9FqpAa8VnCwv
'' SIG '' SRvbIrvD/niUUcOGsYKIXfA9tFGheTMrLnu53CAJE3Hr
'' SIG '' ahlbz+ilMFcsiUk/uc9/yb8+ImhjU5q9aXSsxR08f5Lg
'' SIG '' w7wc2AR1MIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1
'' SIG '' b5VQCDANBgkqhkiG9w0BAQsFADBlMQswCQYDVQQGEwJV
'' SIG '' UzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQL
'' SIG '' ExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdp
'' SIG '' Q2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMTMxMDIy
'' SIG '' MTIwMDAwWhcNMjgxMDIyMTIwMDAwWjByMQswCQYDVQQG
'' SIG '' EwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
'' SIG '' VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhE
'' SIG '' aWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWdu
'' SIG '' aW5nIENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIB
'' SIG '' CgKCAQEA+NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/
'' SIG '' 5aid2zLXcep2nQUut4/6kkPApfmJ1DcZ17aq8JyGpdgl
'' SIG '' rA55KDp+6dFn08b7KSfH03sjlOSRI5aQd4L5oYQjZhJU
'' SIG '' M1B0sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxKhwjf
'' SIG '' DPXiTWAYvqrEsq5wMWYzcT6scKKrzn/pfMuSoeU7MRzP
'' SIG '' 6vIK5Fe7SrXpdOYr/mzLfnQ5Ng2Q7+S1TqSp6moKq4Tz
'' SIG '' rGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi6CxR93O8
'' SIG '' vYWxYoNzQYIH5DiLanMg0A9kczyen6Yzqf0Z3yWT0QID
'' SIG '' AQABo4IBzTCCAckwEgYDVR0TAQH/BAgwBgEB/wIBADAO
'' SIG '' BgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUH
'' SIG '' AwMweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhho
'' SIG '' dHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUH
'' SIG '' MAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9E
'' SIG '' aWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1Ud
'' SIG '' HwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0
'' SIG '' LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmww
'' SIG '' OqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9E
'' SIG '' aWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwTwYDVR0g
'' SIG '' BEgwRjA4BgpghkgBhv1sAAIEMCowKAYIKwYBBQUHAgEW
'' SIG '' HGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCgYI
'' SIG '' YIZIAYb9bAMwHQYDVR0OBBYEFFrEuXsqCqOl6nEDwGD5
'' SIG '' LfZldQ5YMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
'' SIG '' IZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+7A1aJLPz
'' SIG '' ItEVyCx8JSl2qB1dHC06GsTvMGHXfgtg/cM9D8Svi/3v
'' SIG '' Kt8gVTew4fbRknUPUbRupY5a4l4kgU4QpO4/cY5jDhNL
'' SIG '' rddfRHnzNhQGivecRk5c/5CxGwcOkRX7uq+1UcKNJK4k
'' SIG '' xscnKqEpKBo6cSgCPC6Ro8AlEeKcFEehemhor5unXCBc
'' SIG '' 2XGxDI+7qPjFEmifz0DLQESlE/DmZAwlCEIysjaKJAL+
'' SIG '' L3J+HNdJRZboWR3p+nRka7LrZkPas7CM1ekN3fYBIM6Z
'' SIG '' MWM9CBoYs4GbT8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiF
'' SIG '' LpKR6mhsRDKyZqHnGKSaZFHvMIIFWDCCBECgAwIBAgIQ
'' SIG '' BOIBAMdcvHSgrnoelKcXbjANBgkqhkiG9w0BAQsFADBy
'' SIG '' MQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQg
'' SIG '' SW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEw
'' SIG '' LwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQg
'' SIG '' Q29kZSBTaWduaW5nIENBMB4XDTE4MDIxNjAwMDAwMFoX
'' SIG '' DTIxMDIyNDEyMDAwMFowgZQxCzAJBgNVBAYTAlVTMQsw
'' SIG '' CQYDVQQIEwJDQTETMBEGA1UEBxMKRW1lcnl2aWxsZTEU
'' SIG '' MBIGA1UEChMLVGFuaXVtIEluYy4xEzARBgNVBAsTClBv
'' SIG '' d2Vyc2hlbGwxFDASBgNVBAMTC1Rhbml1bSBJbmMuMSIw
'' SIG '' IAYJKoZIhvcNAQkBFhNzZWN1cml0eUB0YW5pdW0uY29t
'' SIG '' MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
'' SIG '' tETqLMiVcPUluh58k/cLjGY8nniDDH4NIDALeI0SjrcZ
'' SIG '' fyxGtFwbZMEMI9Vl1witQmRC59mRkASxifW1rtU0OIqH
'' SIG '' Mou5DTUurIeQymoeLlkbZtlsK1xGZTFMxDoT2uLdOS2+
'' SIG '' roKuqGod4PJOjLk69OQhGjrJ6gaPfgyYe7C8jGASPSnC
'' SIG '' j2iuQe2oY74Dt1FhVtkcx5KDS2V9zp7kagu5IeHWE13B
'' SIG '' CrvTX8d9MmDHWVZCFk5fBDAs76TPeXQmaIFmu0erm/D7
'' SIG '' CH+ni1a35XZod6MOeFVAOt5voUIT7pYXk3/WAA24EZdr
'' SIG '' 7VwwJebomuhGJMLLIpiawY8mEaCKEODdWwIDAQABo4IB
'' SIG '' xTCCAcEwHwYDVR0jBBgwFoAUWsS5eyoKo6XqcQPAYPkt
'' SIG '' 9mV1DlgwHQYDVR0OBBYEFPo9cXkLStRV4KD3xoBdkzV4
'' SIG '' ukA3MA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggr
'' SIG '' BgEFBQcDAzB3BgNVHR8EcDBuMDWgM6Axhi9odHRwOi8v
'' SIG '' Y3JsMy5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVkLWNz
'' SIG '' LWcxLmNybDA1oDOgMYYvaHR0cDovL2NybDQuZGlnaWNl
'' SIG '' cnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwTAYD
'' SIG '' VR0gBEUwQzA3BglghkgBhv1sAwEwKjAoBggrBgEFBQcC
'' SIG '' ARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAI
'' SIG '' BgZngQwBBAEwgYQGCCsGAQUFBwEBBHgwdjAkBggrBgEF
'' SIG '' BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4G
'' SIG '' CCsGAQUFBzAChkJodHRwOi8vY2FjZXJ0cy5kaWdpY2Vy
'' SIG '' dC5jb20vRGlnaUNlcnRTSEEyQXNzdXJlZElEQ29kZVNp
'' SIG '' Z25pbmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG
'' SIG '' 9w0BAQsFAAOCAQEAg2VYTYM3kfyNJKeeAJEForMPHxQq
'' SIG '' tQDWumyFjeNghzpI5+Ppq7A6SHRT5My31Nl2IxRZdHs/
'' SIG '' MBnOdEmuRyQZ6QMG8PuEFSZ1IPuo3hGdUF2Im5+By2w0
'' SIG '' cU3b9RoMOAeV93z4g3+58BTLCx3UQx8bduV5ibgfrWtZ
'' SIG '' 4ec7gG1F8YxOjqLkEAbvr6K38TCzZ3XSvW23jMSYBRCw
'' SIG '' 4Q8loXJWUGjNW534oYF1+VgvAEj2Vf6gLZuJTqmiBpdm
'' SIG '' QfTTkq9oIi3RMuNtvmfCbstpxOLfc6g+0/v0mD0oESYu
'' SIG '' 6M8vUjDSVMNNgN8yLY6n6B92AshDaQYCJTANmZ395MGB
'' SIG '' S33tRjGCBDcwggQzAgEBMIGGMHIxCzAJBgNVBAYTAlVT
'' SIG '' MRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsT
'' SIG '' EHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lD
'' SIG '' ZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcg
'' SIG '' Q0ECEATiAQDHXLx0oK56HpSnF24wCQYFKw4DAhoFAKB4
'' SIG '' MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZI
'' SIG '' hvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIB
'' SIG '' CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYE
'' SIG '' FATu7hF24jyDahqDKMyouTKjJ8+HMA0GCSqGSIb3DQEB
'' SIG '' AQUABIIBAIo7qxucTzuuBufID7RP1VEaGuY61D8GuDN3
'' SIG '' COBbsYP++WVle13Do/9eeKgCVjdOTeP/S5AtS7BLRc82
'' SIG '' GUvw+DzFxm+289xiWg+u4H31lqkLg8ER757UWAbLK1nH
'' SIG '' lkCK61oQbck+ZYjCC6q6wBIs3x+e7yw8q1UC9LTib8iC
'' SIG '' 5fWLn9mlnjCD1bWMYiMzQqNlwbvvOKeVsILpUwPYUcwx
'' SIG '' MwROVf4xxaK2GrZ1glannxfGHNlfNGXOlfVhAdCXOd8E
'' SIG '' uwV2v5t5jQGgO7Kpy83vVI4Hx9Gyw9Y5cci6QmtNZs1E
'' SIG '' Bgv849jXSQaW3IcIEEwpdWzNvOaugN2iqycnXoWljveh
'' SIG '' ggILMIICBwYJKoZIhvcNAQkGMYIB+DCCAfQCAQEwcjBe
'' SIG '' MQswCQYDVQQGEwJVUzEdMBsGA1UEChMUU3ltYW50ZWMg
'' SIG '' Q29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFudGVjIFRp
'' SIG '' bWUgU3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMgIQDs/0
'' SIG '' OMj+vzVuBNhqmBsaUDAJBgUrDgMCGgUAoF0wGAYJKoZI
'' SIG '' hvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
'' SIG '' DxcNMTgxMDIzMTYzMDU5WjAjBgkqhkiG9w0BCQQxFgQU
'' SIG '' NWHK6A7PPAxqQxMp8afqxld1mC8wDQYJKoZIhvcNAQEB
'' SIG '' BQAEggEARuFbnMnBUzQJ+sCNPI+gRiwlEQ6RchNadPue
'' SIG '' oMegOma9Gh9HJer70PmXmKd0NswJuSC2QPjF7gL2Pv3H
'' SIG '' FQdSPgxuGRYeCsQUMZYaVI0S02b2yvAM73+HZzLtUxCK
'' SIG '' x2+J5HtPDOZ7Mwb5sl88RHZpza3vk53KMHY4h0Ovv5p4
'' SIG '' irosU4Wu0sMsb85kwjlaExcOQT/qutlVUbkCGUD4ND1Y
'' SIG '' zw0jiShWjhvCCWmWegXhwWAXxnabgV/gm1GAP3ZZOfyB
'' SIG '' y1QUq55kE1XFV3WhtuBkeOs4YOhClvrQThACls3CBwMZ
'' SIG '' mkbpR+0RrG/N5MakSA2uY+xy8n9Y9tHZ8WCghS41QQ==
'' SIG '' End signature block
