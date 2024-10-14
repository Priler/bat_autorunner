use {
    std::{
        env,
        io,
    },
    winresource::WindowsResource,
};

fn main() -> io::Result<()> {
    if cfg!(target_os = "windows") {
        let mut res = winresource::WindowsResource::new();
        res
            // This path can be absolute, or relative to your crate root.
            .set_icon("icon.ico")
            .set("InternalName", "Автозапуск БАТника как службы by Priler");

        res.compile()?;
    }

    Ok(())
}