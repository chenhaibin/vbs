@echo off

if "%1"=="" goto error
if "%2"=="" goto error

start cscript afeventlog.vbs %1 %2 ecom0201
start cscript afeventlog.vbs %1 %2 ecom0202
start cscript afeventlog.vbs %1 %2 ecom0203
start cscript afeventlog.vbs %1 %2 ecom0204
start cscript afeventlog.vbs %1 %2 ecom0205
start cscript afeventlog.vbs %1 %2 ecom0206
start cscript afeventlog.vbs %1 %2 ecom0207
start cscript afeventlog.vbs %1 %2 ecom0208
start cscript afeventlog.vbs %1 %2 ecom0209
start cscript afeventlog.vbs %1 %2 ecom0210
start cscript afeventlog.vbs %1 %2 ecom0218
start cscript afeventlog.vbs %1 %2 ecom0219
start cscript afeventlog.vbs %1 %2 ecom0220
start cscript afeventlog.vbs %1 %2 ecom0221
start cscript afeventlog.vbs %1 %2 ecom0222
start cscript afeventlog.vbs %1 %2 ecom0225
start cscript afeventlog.vbs %1 %2 ecom0226
start cscript afeventlog.vbs %1 %2 ecom0227

start cscript afeventlog.vbs %1 %2 ecom5201
start cscript afeventlog.vbs %1 %2 ecom5202
start cscript afeventlog.vbs %1 %2 ecom5203
start cscript afeventlog.vbs %1 %2 ecom5204
start cscript afeventlog.vbs %1 %2 ecom5205
start cscript afeventlog.vbs %1 %2 ecom5206
start cscript afeventlog.vbs %1 %2 ecom5207
start cscript afeventlog.vbs %1 %2 ecom5208
start cscript afeventlog.vbs %1 %2 ecom5209
start cscript afeventlog.vbs %1 %2 ecom5210
start cscript afeventlog.vbs %1 %2 ecom5218
start cscript afeventlog.vbs %1 %2 ecom5219
start cscript afeventlog.vbs %1 %2 ecom5220
start cscript afeventlog.vbs %1 %2 ecom5221
start cscript afeventlog.vbs %1 %2 ecom5222
start cscript afeventlog.vbs %1 %2 ecom5225
start cscript afeventlog.vbs %1 %2 ecom5226
start cscript afeventlog.vbs %1 %2 ecom5227

goto end

:error
echo.
echo  No Arguments specified

:end 