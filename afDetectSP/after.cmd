@echo off
REM /****************************************************************************/
REM /*                                                                          */
REM /*  PROJECT:  MS03-043 Patch	                                        */
REM /*  PLATFORM: E-Commerce                                                    */
REM /*                                                                          */
REM /* AUTHOR                                                                   */
REM /*  James Taylor                                                            */
REM /*                                                                          */
REM /* DATE                                                                     */
REM /*  20th October 2003	- (1.0) Created  				*/
REM /*                                                                          */
REM /****************************************************************************/

REM /****************************************************************************/
REM /* Validate parameters							*/
REM /* %1 - Installation files location						*/
REM /****************************************************************************/

If %1#==?#  GOTO ERR_PARAM
If %1#==#   GOTO ERR_PARAM

REM /****************************************************************************/
REM /* Set source directory using parameter 1 from command string               */
REM /****************************************************************************/

set SOURCE_DIR=%1

REM /****************************************************************************/
REM /* Run Detection Script with patch install			                */
REM /****************************************************************************/
Echo Running Detection Script which will then install correct patch

Cscript %1/ms03-043/detect-install.vbs

IF NOT %ERRORLEVEL% == 0 GOTO ERR_SCRIPT
GOTO EOF

:ERR_SCRIPT
REM ********************************************************************************
REM /* Error running script							
REM ********************************************************************************

ECHO Script failed during detection & patch install
GOTO EOF

:EOF