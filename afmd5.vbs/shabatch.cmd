sha1sum %1 >> %2.info
if %errorlevel% 1 goto error

goto end

:error
echo sha1sum errored on file %1 >> %2.info
goto end

:end