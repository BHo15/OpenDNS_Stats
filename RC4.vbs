'http://www.visualbasicscript.com/Encrypt-a-string-m48428.aspx

 Function fCrypt(sPlainText, sPassword)
 'This function will encrypt or decrypt a string using the RSA's RC4 algorithm.
  Dim aBox(255), aKey(255), sTemp, a, b, c, i, j, k, iCipherBy, sTempswap, iLength, sO
  i = 0:j = 0:b = 0
  iLength = Len(sPassword)
   For a = 0 To 255
    aKey(a) = Asc(Mid(sPassword, (a Mod iLength)+1, 1))
    aBox(a) = a
   Next
   For a = 0 To 255
    b = (b + aBox(a) + aKey(a)) Mod 256
    sTempswap = aBox(a)
    aBox(a) = aBox(b)
    aBox(b) = sTempswap
   Next
   For c = 1 To Len(sPlainText)
    i = (i + 1) Mod 256
    j = (j + aBox(i)) Mod 256
    sTemp = aBox(i)
    aBox(i) = aBox(j)
    aBox(j) = sTemp
    k = aBox((aBox(i) + aBox(j)) Mod 256)
    iCipherBy = Asc(Mid(sPlainText, c, 1)) Xor k
    sO = sO & Chr(iCipherBy)
   Next
  fCrypt = sO
 End Function
 