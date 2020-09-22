<div align="center">

## ASCII\_TO\_EBCDIC


</div>

### Description

Convert ASCII strings into EBCDIC code to upload into an IBM mainframe. This code may be used also as a basic encrypting method. Both ASCII to EBCDIC and EBCDIC to ASCII are included in this code.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Aldo Vargas](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/aldo-vargas.md)
**Level**          |Intermediate
**User Rating**    |5.0 (5 globes from 1 user)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/aldo-vargas-ascii-to-ebcdic__1-5902/archive/master.zip)





### Source Code

```
Function ascii_to_ebcdic(ByVal buffer As String) As String
 Dim ascii As Variant
 Dim i As Long, bufferlen As Long
 ascii = Array( _
  &H0, &H1, &H2, &H3, &H37, &H2D, &H2E, &H2F, &H16, &H5, &H25, &HB, &HC, &HD, &HE, &HF, _
  &H10, &H11, &H12, &H13, &H3C, &H3D, &H32, &H26, &H18, &H19, &H3F, &H27, &H1C, &H1D, &H1E, &H1F, _
  &H40, &H4F, &H7F, &H7B, &H5B, &H6C, &H50, &H7D, &H4D, &H5D, &H5C, &H4E, &H6B, &H60, &H4B, &H61, _
  &HF0, &HF1, &HF2, &HF3, &HF4, &HF5, &HF6, &HF7, &HF8, &HF9, &H7A, &H5E, &H4C, &H7E, &H6E, &H6F, _
  &H7C, &HC1, &HC2, &HC3, &HC4, &HC5, &HC6, &HC7, &HC8, &HC9, &HD1, &HD2, &HD3, &HD4, &HD5, &HD6, _
  &HD7, &HD8, &HD9, &HE2, &HE3, &HE4, &HE5, &HE6, &HE7, &HE8, &HE9, &H4A, &HE0, &H5A, &H5F, &H6D, _
  &H79, &H81, &H82, &H83, &H84, &H85, &H86, &H87, &H88, &H89, &H91, &H92, &H93, &H94, &H95, &H96, _
  &H97, &H98, &H99, &HA2, &HA3, &HA4, &HA5, &HA6, &HA7, &HA8, &HA9, &HC0, &H6A, &HD0, &HA1, &H7, _
  &H20, &H21, &H22, &H23, &H24, &H15, &H6, &H17, &H28, &H29, &H2A, &H2B, &H2C, &H9, &HA, &H1B, _
  &H30, &H31, &H1A, &H33, &H34, &H35, &H36, &H8, &H38, &H39, &H3A, &H3B, &H4, &H14, &H3E, &HE1, _
  &H41, &H42, &H43, &H44, &H45, &H46, &H47, &H48, &H49, &H51, &H52, &H53, &H54, &H55, &H56, &H57, _
  &H58, &H59, &H62, &H63, &H64, &H65, &H66, &H67, &H68, &H69, &H70, &H71, &H72, &H73, &H74, &H75, _
  &H76, &H77, &H78, &H80, &H8A, &H8B, &H8C, &H8D, &H8E, &H8F, &H90, &H9A, &H9B, &H9C, &H9D, &H9E, _
  &H9F, &HA0, &HAA, &HAB, &HAC, &HAD, &HAE, &HAF, &HB0, &HB1, &HB2, &HB3, &HB4, &HB5, &HB6, &HB7, _
  &HB8, &HB9, &HBA, &HBB, &HBC, &HBD, &HBE, &HBF, &HCA, &HCB, &HCC, &HCD, &HCE, &HCF, &HDA, &HDB, _
  &HDC, &HDD, &HDE, &HDF, &HEA, &HEB, &HEC, &HED, &HEE, &HEF, &HFA, &HFB, &HFC, &HFD, &HFE, &HFF)
 bufferlen = Len(buffer)
 For i = 1 To bufferlen
  Mid$(buffer, i, 1) = Chr$(ascii(Asc(Mid$(buffer, i, 1))))
 Next
 ascii_to_ebcdic = buffer
End Function
Function ebcdic_to_ascii(ByVal buffer As String) As String
 Dim ebcdic As Variant
 Dim i As Long, bufferlen As Long
 ebcdic = Array( _
  &H0, &H1, &H2, &H3, &H9C, &H9, &H86, &H7F, &H97, &H8D, &H8E, &HB, &HC, &HD, &HE, &HF, _
  &H10, &H11, &H12, &H13, &H9D, &H85, &H8, &H87, &H18, &H19, &H92, &H8F, &H1C, &H1D, &H1E, &H1F, _
  &H80, &H81, &H82, &H83, &H84, &HA, &H17, &H1B, &H88, &H89, &H8A, &H8B, &H8C, &H5, &H6, &H7, _
  &H90, &H91, &H16, &H93, &H94, &H95, &H96, &H4, &H98, &H99, &H9A, &H9B, &H14, &H15, &H9E, &H1A, _
  &H20, &HA0, &HA1, &HA2, &HA3, &HA4, &HA5, &HA6, &HA7, &HA8, &H5B, &H2E, &H3C, &H28, &H2B, &H21, _
  &H26, &HA9, &HAA, &HAB, &HAC, &HAD, &HAE, &HAF, &HB0, &HB1, &H5D, &H24, &H2A, &H29, &H3B, &H5E, _
  &H2D, &H2F, &HB2, &HB3, &HB4, &HB5, &HB6, &HB7, &HB8, &HB9, &H7C, &H2C, &H25, &H5F, &H3E, &H3F, _
  &HBA, &HBB, &HBC, &HBD, &HBE, &HBF, &HC0, &HC1, &HC2, &H60, &H3A, &H23, &H40, &H27, &H3D, &H22, _
  &HC3, &H61, &H62, &H63, &H64, &H65, &H66, &H67, &H68, &H69, &HC4, &HC5, &HC6, &HC7, &HC8, &HC9, _
  &HCA, &H6A, &H6B, &H6C, &H6D, &H6E, &H6F, &H70, &H71, &H72, &HCB, &HCC, &HCD, &HCE, &HCF, &HD0, _
  &HD1, &H7E, &H73, &H74, &H75, &H76, &H77, &H78, &H79, &H7A, &HD2, &HD3, &HD4, &HD5, &HD6, &HD7, _
  &HD8, &HD9, &HDA, &HDB, &HDC, &HDD, &HDE, &HDF, &HE0, &HE1, &HE2, &HE3, &HE4, &HE5, &HE6, &HE7, _
  &H7B, &H41, &H42, &H43, &H44, &H45, &H46, &H47, &H48, &H49, &HE8, &HE9, &HEA, &HEB, &HEC, &HED, _
  &H7D, &H4A, &H4B, &H4C, &H4D, &H4E, &H4F, &H50, &H51, &H52, &HEE, &HEF, &HF0, &HF1, &HF2, &HF3, _
  &H5C, &H9F, &H53, &H54, &H55, &H56, &H57, &H58, &H59, &H5A, &HF4, &HF5, &HF6, &HF7, &HF8, &HF9, _
  &H30, &H31, &H32, &H33, &H34, &H35, &H36, &H37, &H38, &H39, &HFA, &HFB, &HFC, &HFD, &HFE, &HFF)
 bufferlen = Len(buffer)
 For i = 1 To bufferlen
  Mid$(buffer, i, 1) = Chr$(ebcdic(Asc(Mid$(buffer, i, 1))))
 Next
 ebcdic_to_ascii = buffer
End Function
```

