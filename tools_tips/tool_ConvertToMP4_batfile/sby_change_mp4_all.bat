
@echo off
rem ファイル名の一覧を取得


for %%A in (*.flv) do (
        echo %%A
        call sby_change_mp4.bat "%%A"
)
for %%A in (*.mkv) do (
        echo %%A
        call sby_change_mp4.bat "%%A"
)
for %%A in (*.f4v) do (
        echo %%A
        call sby_change_mp4.bat "%%A"
)
