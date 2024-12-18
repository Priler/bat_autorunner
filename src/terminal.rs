use std::io::{self, stdout, Write};
use crossterm::{
    cursor,
    terminal::{self, Clear, ClearType},
    execute,
};

pub fn init(stdout: &mut impl Write) -> io::Result<()> {
    terminal::enable_raw_mode()?;
    execute!(
        stdout,
        terminal::EnterAlternateScreen,
        cursor::Hide,
        Clear(ClearType::All),
        cursor::MoveTo(0, 0)
    )?;
    Ok(())
}

pub fn get_size() -> io::Result<(u16, u16)> {
    terminal::size()
}

pub fn cleanup_terminal() -> io::Result<()> {
    let mut stdout = stdout();
    terminal::disable_raw_mode()?;
    execute!(
        stdout,
        Clear(ClearType::All),
        cursor::Show,
        terminal::LeaveAlternateScreen
    )?;
    stdout.flush()?;
    Ok(())
}

pub fn setup_terminal_cleanup() {
    let original_hook = std::panic::take_hook();
    std::panic::set_hook(Box::new(move |panic_info| {
        let _ = cleanup_terminal();
        original_hook(panic_info);
    }));

    ctrlc::set_handler(move || {
        let _ = cleanup_terminal();
        std::process::exit(0);
    }).expect("Error setting Ctrl-C handler");
}

pub fn cleanup_and_exit(stdout: &mut impl Write) -> io::Result<()> {
    //execute!(
    //    stdout,
    //    cursor::MoveTo(0, 0),
    //    Clear(ClearType::All)
    //)?;
    // cleanup_terminal()?;
    println!("\nГотово!\nВы можете закрыть это окно.");
    let mut input = String::new();
    io::stdin().read_line(&mut input)?;
    Ok(())
}