# office
```
.\setup.exe /configure filename.xml
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
