mod terminal;
mod service;
mod ui;
mod utils;

use std::io::{self, stdout};
use terminal::{cleanup_terminal, setup_terminal_cleanup};
use ui::{print_welcome_message, render_options};
use utils::get_options;

fn main() -> io::Result<()> {
    setup_terminal_cleanup();
    let mut stdout = stdout();

    // Initialize terminal
    terminal::init(&mut stdout)?;

    // Get terminal size early
    let (_, term_height) = terminal::get_size()?;

    // Print welcome messages and get the starting line
    let current_line = print_welcome_message();

    // Get the list of .bat files and add options
    let options = get_options();
    if options.is_empty() {
        println!("Не найдено ни одного .bat файла в текущей директории.");
        cleanup_terminal()?;
        return Ok(());
    }

    ui::run_main_loop(
        &mut stdout,
        &options,
        current_line,
        term_height,
    )?;

    terminal::cleanup_and_exit(&mut stdout)?;
    Ok(())
}