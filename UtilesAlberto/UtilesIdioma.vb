Imports System.Threading
Imports System.Globalization

Partial Public Class Utiles
    Public Const IdiomaES = "es-ES"
    Public Shared Sub CultureInfo_SetSpain()
        Try
            '' Poner la configuración regional en España. Para las conversiones de . decimal en ,
            Dim ConfReg As New System.Globalization.CultureInfo(IdiomaES, True)
            System.Threading.Thread.CurrentThread.CurrentCulture = ConfReg
            System.Threading.Thread.CurrentThread.CurrentUICulture = ConfReg
            ConfReg = Nothing
        Catch ex As Exception
            '' No hacemos nada.
        End Try
    End Sub
    Public Shared Sub CultureInfo_Set(strIdioma As String)
        Try
            '' Poner la configuración regional en strIdioma.
            Dim ConfReg As New System.Globalization.CultureInfo(strIdioma, True)
            System.Threading.Thread.CurrentThread.CurrentCulture = ConfReg
            System.Threading.Thread.CurrentThread.CurrentUICulture = ConfReg
            ConfReg = Nothing
        Catch ex As Exception
            '' No hacemos nada.
        End Try
    End Sub
    '
    Public Shared Function CulturesInfo_Read(Optional mensaje As Boolean = False) As String
        Dim resultado As String = ""
        For Each ci As CultureInfo In CultureInfo.GetCultures(CultureTypes.AllCultures)
            resultado &= ci.EnglishName + " - " + ci.Name + vbCrLf
        Next
        '
        If mensaje = True Then
            MsgBox(resultado)
        End If
        Return resultado
    End Function
End Class


'
'af-ZA		'Afrikaans - South Africa
'ar         'Arabic
'ar-AE		'Arabic - United Arab Emirates
'ar-BH		'Arabic - Bahrain
'ar-DZ		'Arabic - Algeria
'ar-EG		'Arabic - Egypt
'ar-IQ		'Arabic - Iraq
'ar-JO		'Arabic - Jordan
'ar-KW		'Arabic - Kuwait
'ar-LB		'Arabic - Lebanon
'ar-LY		'Arabic - Libya
'ar-MA		'Arabic - Morocco
'ar-OM		'Arabic - Oman
'ar-QA		'Arabic - Qatar
'ar-SA		'Arabic - Saudi Arabia
'ar-SY		'Arabic - Syria
'ar-TN		'Arabic - Tunisia
'ar-YE		'Arabic - Yemen
'be-BY		'Belarusian - Belarus
'bg         'Bulgarian
'bg-BG		'Bulgarian - Bulgaria
'ca-ES		'Catalan - Catalan
'cs         'Czech
'cs-CZ		'Czech - Czech Republic
'Cy-az-AZ	'Azeri (Cyrillic) - Azerbaijan
'Cy-sr-SP	'Serbian (Cyrillic) - Serbia
'Cy-uz-UZ	'Uzbek (Cyrillic) - Uzbekistan
'da         'Danish
'da-DK		'Danish - Denmark
'de         'German
'de-AT		'German - Austria
'de-CH		'German - Switzerland
'de-DE		'German - Germany
'de-LI		'German - Liechtenstein
'de-LU		'German - Luxembourg
'div-MV		'Dhivehi - Maldives
'el         'Greek
'el-GR		'Greek - Greece
'en         'English
'en-AU		'English - Australia
'en-BZ		'English - Belize
'en-CA		'English - Canada
'en-CB		'English - Caribbean
'en-GB		'English - United Kingdom
'en-IE		'English - Ireland
'en-JM		'English - Jamaica
'en-NZ		'English - New Zealand
'en-PH		'English - Philippines
'en-TT		'English - Trinidad and Tobago
'en-US		'English - United States
'en-ZA		'English - South Africa
'en-ZW		'English - Zimbabwe
'es         'Spanish
'es-AR		'Spanish - Argentina
'es-BO		'Spanish - Bolivia
'es-CL		'Spanish - Chile
'es-CO		'Spanish - Colombia
'es-CR		'Spanish - Costa Rica
'es-Do		'Spanish - Dominican Republic
'es-EC		'Spanish - Ecuador
'es-ES		'Spanish - Spain
'es-GT		'Spanish - Guatemala
'es-HN		'Spanish - Honduras
'es-MX		'Spanish - Mexico
'es-NI		'Spanish - Nicaragua
'es-PA		'Spanish - Panama
'es-PE		'Spanish - Peru
'es-PR		'Spanish - Puerto Rico
'es-PY		'Spanish - Paraguay
'es-SV		'Spanish - El Salvador
'es-UY		'Spanish - Uruguay
'es-VE		'Spanish - Venezuela
'et-EE		'Estonian - Estonia
'eu-ES		'Basque - Basque
'fa-IR		'Farsi - Iran
'fi         'Finnish
'fi-FI		'Finnish - Finland
'fo-FO		'Faroese - Faroe Islands
'fr-BE		'French - Belgium
'fr-CA		'French - Canada
'fr-CH		'French - Switzerland
'fr-FR		'French - France
'fr-LU		'French - Luxembourg
'fr-MC		'French - Monaco
'gl-ES		'Galician - Galician
'gu-In		'Gujarati - India
'he-IL		'Hebrew - Israel
'hi-In		'Hindi - India
'hr-HR		'Croatian - Croatia
'hu-HU		'Hungarian - Hungary
'hy-AM		'Armenian - Armenia
'id-ID		'Indonesian - Indonesia
'Is-Is		'Icelandic - Iceland
'it-CH		'Italian - Switzerland
'it-IT		'Italian
'it-IT		'Italian - Italy
'ja-JP		'Japanese - Japan
'ka-GE		'Georgian - Georgia
'kk-KZ		'Kazakh - Kazakhstan
'kn-IN		'Kannada - India
'ko-KR		'Korean - Korea
'kok-IN		'Konkani - India
'ky-KZ		'Kyrgyz - Kazakhstan
'Lt-az-AZ	'Azeri (Latin) - Azerbaijan
'lt-LT		'Lithuanian - Lithuania
'Lt-sr-SP	'Serbian (Latin) - Serbia
'Lt-uz-UZ	'Uzbek (Latin) - Uzbekistan
'lv-LV		'Latvian - Latvia
'mk-MK		'Macedonian (FYROM)
'mn-MN		'Mongolian - Mongolia
'mr-IN		'Marathi - India
'ms-BN		'Malay - Brunei
'ms-MY		'Malay - Malaysia
'nb-NO		'Norwegian (Bokmål) - Norway
'nl-BE		'Dutch - Belgium
'nl-NL		'Dutch - The Netherlands
'nn-NO		'Norwegian (Nynorsk) - Norway
'pa-IN		'Punjabi - India
'pl-PL		'Polish - Poland
'pt-BR		'Portuguese - Brazil
'pt-PT		'Portuguese - Portugal
'ro-RO		'Romanian - Romania
'ru-RU		'Russian - Russia
'sa-IN		'Sanskrit - India
'sk-SK		'Slovak - Slovakia
'sl-SI		'Slovenian - Slovenia
'sq-AL		'Albanian - Albania
'sv-FI		'Swedish - Finland
'sv-SE		'Swedish - Sweden
'sw-KE		'Swahili - Kenya
'syr-SY		'Syriac - Syria
'ta-IN		'Tamil - India
'te-IN		'Telugu - India
'th-TH		'Thai - Thailand
'tr-TR		'Turkish - Turkey
'tt-RU		'Tatar - Russia
'uk-UA		'Ukrainian - Ukraine
'ur-PK		'Urdu - Pakistan
'vi-VN		'Vietnamese - Vietnam
'zh		    'Chinese
'zh-CHS		'Chinese (Simplified)
'zh-CHT		'Chinese (Traditional)
'zh-CN		'Chinese - China
'zh-HK		'Chinese - Hong Kong SAR
'zh-MO		'Chinese - Macau SAR
'zh-SG		'Chinese - Singapore
'zh-TW		'Chinese - Taiwan
