# office
```
.\setup.exe /configure filename.xml
```

```
Office Version	 Office Product          Generic Key                     Key Type
v16.0 (2024)     ProPlus2024Volume       4YV2J-VNG7W-YGTP3-443TK-TF8CP   MAK-AE1
v16.0 (2024)     VisioPro2024Volume      GBNHB-B2G3Q-G42YB-3MFC2-7CJCX   MAK-AE
v16.0 (2024)     ProjectPro2024Volume    WNFMR-HK4R7-7FJVM-VQ3JC-76HF6   MAK-AE1

```

```
pushd "C:\Program Files\Microsoft Office\Office16"

cscript ospp.vbs /act
cscript ospp.vbs /inpkey:value
cscript ospp.vbs /unpkey:value
cscript ospp.vbs /sethst:value
cscript ospp.vbs /setprt:value
cscript ospp.vbs /remhst

cscript ospp.vbs /dstatus
cscript ospp.vbs /dstatusall

```
