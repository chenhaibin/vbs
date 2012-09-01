@echo off

if "%1"=="" goto error
if "%2"=="" goto error

start cscript afeventlog.vbs %1 %2 ebus2201
start cscript afeventlog.vbs %1 %2 ebus2202
start cscript afeventlog.vbs %1 %2 ebus2203
start cscript afeventlog.vbs %1 %2 ebus2205
start cscript afeventlog.vbs %1 %2 ebus2207
start cscript afeventlog.vbs %1 %2 ebus2208
start cscript afeventlog.vbs %1 %2 ebus2209
start cscript afeventlog.vbs %1 %2 ebus2210
start cscript afeventlog.vbs %1 %2 ebus2211
start cscript afeventlog.vbs %1 %2 ebus2212
start cscript afeventlog.vbs %1 %2 ebus2213
start cscript afeventlog.vbs %1 %2 ebus2214
start cscript afeventlog.vbs %1 %2 ebus2215
start cscript afeventlog.vbs %1 %2 ebus2216
start cscript afeventlog.vbs %1 %2 ebus2217
start cscript afeventlog.vbs %1 %2 ebus2218
start cscript afeventlog.vbs %1 %2 ebus2219
start cscript afeventlog.vbs %1 %2 ebus2220
start cscript afeventlog.vbs %1 %2 ebus2225
start cscript afeventlog.vbs %1 %2 ebus2226
start cscript afeventlog.vbs %1 %2 ebus2227
start cscript afeventlog.vbs %1 %2 ebus2228
start cscript afeventlog.vbs %1 %2 ebus2229
start cscript afeventlog.vbs %1 %2 ebus2230
start cscript afeventlog.vbs %1 %2 ebus2231
start cscript afeventlog.vbs %1 %2 ebus2232
start cscript afeventlog.vbs %1 %2 ebus2233
start cscript afeventlog.vbs %1 %2 ebus2234
start cscript afeventlog.vbs %1 %2 ebus2235
start cscript afeventlog.vbs %1 %2 ebus2236
start cscript afeventlog.vbs %1 %2 ebus2237
start cscript afeventlog.vbs %1 %2 ebus2103
start cscript afeventlog.vbs %1 %2 ebus2106
start cscript afeventlog.vbs %1 %2 ebus2107

start cscript afeventlog.vbs %1 %2 ebus7201
start cscript afeventlog.vbs %1 %2 ebus7202
start cscript afeventlog.vbs %1 %2 ebus7203
start cscript afeventlog.vbs %1 %2 ebus7205
start cscript afeventlog.vbs %1 %2 ebus7103
::start cscript afeventlog.vbs %1 %2 ebus7106
::start cscript afeventlog.vbs %1 %2 ebus7107

goto end

:error
echo.
echo  No Arguments specified

:end 