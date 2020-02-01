@Echo OFF

FOR /D /R %%# in (*) DO (
    PUSHD "%%#"
    FOR %%@ in ("base*") DO (
        Echo Ren: ".\%%~n#\%%@" "%%~n#%%~x@"
        Ren "%%@" "%%~n#%%~x@"
    )
    POPD
)

Pause&Exit