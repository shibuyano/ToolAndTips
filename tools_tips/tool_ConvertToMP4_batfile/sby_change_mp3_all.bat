
@echo off
rem ファイル名の一覧を取得


for %%A in (*.mkv) do (
        echo %%A
        call sby_change_mp3.bat "%%A"
        rem call test "%%A"
)
for %%A in (*.webm) do (
        echo %%A
        call sby_change_mp3.bat "%%A"
        rem call test "%%A"
)
for %%A in (*.mp4) do (
        echo %%A
        call sby_change_mp3.bat "%%A"
        rem call test "%%A"
)
