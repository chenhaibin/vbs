
@echo off

if "%1"=="" goto error
if "%2"=="" goto error

eldump -s \\ebus0201 -l application -c # -K -a %1 -b %2 > c:\temp.log
eldump -s \\ebus0202 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0203 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0205 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0207 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0208 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0209 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0210 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0211 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0212 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0213 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0214 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0215 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0216 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0217 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0218 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0219 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0220 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0225 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0226 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0227 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0228 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0229 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0230 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0231 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0232 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0233 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0234 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0235 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0236 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0237 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0103 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0105 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0106 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus0107 -l application -c # -K -a %1 -b %2 >> c:\temp.log

eldump -s \\ebus5201 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5202 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5203 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5205 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5207 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5208 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5209 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5210 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5211 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5212 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5213 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5214 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5215 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5216 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5217 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5218 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5219 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5220 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5225 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5226 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5227 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5228 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5229 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5230 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5231 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5232 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5233 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5234 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5235 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5236 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5237 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5103 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5105 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5106 -l application -c # -K -a %1 -b %2 >> c:\temp.log
eldump -s \\ebus5107 -l application -c # -K -a %1 -b %2 >> c:\temp.log

goto end
:error
echo.
echo Usage: LogRetrieval.cmd [startdatetime] [enddatetime]
echo.
echo 		date/time format is YYYYMMDDHHmm
echo.
goto end

:end