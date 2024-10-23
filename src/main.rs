use crossterm::{
    cursor,
    event::{self, Event, KeyCode},
    execute,
    terminal::{self, ClearType},
    style::{Color, SetForegroundColor, ResetColor},
};
use std::{
    env,
    fs,
    io::{self, stdout, Write},
    process::Command,
    time::{Duration, Instant},
};

fn main() -> crossterm::Result<()> {
    let mut stdout = stdout();
    execute!(stdout, terminal::EnterAlternateScreen, cursor::Hide)?;
    terminal::enable_raw_mode()?;

    // Clear the screen
    execute!(stdout, terminal::Clear(ClearType::All))?;

    // Print welcome messages and get the starting line for options
    let current_line = print_welcome_message();

    // Get terminal size
    let (_, term_height) = terminal::size()?;

    // Get the list of .bat files and add the option to remove the service
    let options = get_options();
    if options.is_empty() {
        println!("Не найдено ни одного .bat файла в текущей директории.");
        cleanup_terminal()?;
        return Ok(());
    }

    let mut current_selection = 0;
    let start_row = current_line;

    // Calculate maximum visible options
    let max_visible_options = (term_height - start_row as u16 - 2) as usize; // Leave 2 lines for messages
    let mut scroll_offset = 0;

    // Render the initial list of options
    render_options(
        &mut stdout,
        &options,
        current_selection,
        start_row,
        scroll_offset,
        max_visible_options,
    )?;

    // Variable to keep track of where to print messages
    let mut message_row = start_row + std::cmp::min(options.len(), max_visible_options);

    // Variable to track the time of the last key event
    let mut last_event_time = Instant::now();

    // Main event loop for user interaction
    loop {
        if event::poll(Duration::from_millis(100))? {
            if let Event::Key(event) = event::read()? {
                let now = Instant::now();
                // Delay to prevent rapid key repeat
                if now.duration_since(last_event_time) > Duration::from_millis(150) {
                    match event.code {
                        KeyCode::Up => {
                            if current_selection > 0 {
                                current_selection -= 1;
                                // Adjust scroll if selection goes above viewport
                                if current_selection < scroll_offset {
                                    scroll_offset = current_selection;
                                }
                                render_options(
                                    &mut stdout,
                                    &options,
                                    current_selection,
                                    start_row,
                                    scroll_offset,
                                    max_visible_options,
                                )?;
                            }
                        }
                        KeyCode::Down => {
                            if current_selection < options.len() - 1 {
                                current_selection += 1;
                                // Adjust scroll if selection goes below viewport
                                if current_selection >= scroll_offset + max_visible_options {
                                    scroll_offset = current_selection - max_visible_options + 1;
                                }
                                render_options(
                                    &mut stdout,
                                    &options,
                                    current_selection,
                                    start_row,
                                    scroll_offset,
                                    max_visible_options,
                                )?;
                            }
                        }
                        KeyCode::PageUp => {
                            if current_selection > 0 {
                                current_selection = current_selection.saturating_sub(max_visible_options);
                                scroll_offset = scroll_offset.saturating_sub(max_visible_options);
                                render_options(
                                    &mut stdout,
                                    &options,
                                    current_selection,
                                    start_row,
                                    scroll_offset,
                                    max_visible_options,
                                )?;
                            }
                        }
                        KeyCode::PageDown => {
                            if current_selection < options.len() - 1 {
                                current_selection = std::cmp::min(
                                    current_selection + max_visible_options,
                                    options.len() - 1,
                                );
                                scroll_offset = std::cmp::min(
                                    scroll_offset + max_visible_options,
                                    options.len().saturating_sub(max_visible_options),
                                );
                                render_options(
                                    &mut stdout,
                                    &options,
                                    current_selection,
                                    start_row,
                                    scroll_offset,
                                    max_visible_options,
                                )?;
                            }
                        }
                        KeyCode::Enter => {
                            // Clear messages below the options
                            execute!(
                                stdout,
                                cursor::MoveTo(0, message_row as u16),
                                terminal::Clear(ClearType::FromCursorDown)
                            )?;
                            stdout.flush()?;

                            if options[current_selection] == "УДАЛИТЬ СЛУЖБУ С АВТОЗАПУСКА" {
                                message_row = handle_service_removal(&mut stdout, message_row)?;
                            } else if options[current_selection] == "ЗАПУСТИТЬ BLOCKCHECK (АВТО-ПОИСК)" {
                                message_row += 1;
                                execute!(stdout, cursor::MoveTo(0, message_row as u16))?;
                                match run_powershell_command(
                                    "Start-Process 'blockcheck.cmd'",
                                ) {
                                    Ok(_) => {
                                        println!("Blockcheck успешно запущен.");
                                        std::process::exit(0);
                                    },
                                    Err(e) => println!("Ошибка при запуске Blockcheck: {}", e),
                                }
                            }
                            else {
                                let selected_file = &options[current_selection];
                                message_row = handle_service_installation(
                                    &mut stdout,
                                    selected_file,
                                    message_row,
                                )?;
                            }
                            break;
                        }
                        KeyCode::Esc => break,
                        _ => {}
                    }
                    last_event_time = now;
                }
            }
        }
    }

    println!("\nГотово! Вы можете закрыть это окно.");
    println!("Нажмите Enter для выхода.");

    let mut _input = String::new();
    io::stdin()
        .read_line(&mut _input)
        .expect("Ошибка при ожидании ввода.");

    Ok(())
}

/// Prints the welcome message and returns the current line number.
fn print_welcome_message() -> usize {
    let mut current_line = 0;
    println!("Добро пожаловать!");
    current_line += 1;
    println!("Эта программа позволяет установить .bat файл как службу с автозапуском.");
    current_line += 1;
    println!("Автор программы: Хауди (Абрахам) https://github.com/Priler");
    current_line += 1;
    println!("Версия: {}", env!("CARGO_PKG_VERSION"));
    current_line += 1;
    println!("===");
    current_line += 1;
    println!(
        "Используя СТРЕЛКИ на клавиатуре, выберите .bat файл из списка \
        для установки службы 'discordfix_zapret' или выберите \
        'УДАЛИТЬ СЛУЖБУ С АВТОЗАПУСКА' или 'ЗАПУСТИТЬ BLOCKCHECK (АВТО-ПОИСК)'.\n"
    );
    current_line += 2; // Account for the extra newline
    println!("Для выбора нажмите ENTER.");
    current_line += 1;
    current_line
}

/// Retrieves a sorted list of .bat files in the current directory and adds the remove service option.
fn get_options() -> Vec<String> {
    let current_dir = env::current_dir().expect("Не удалось получить текущую директорию");
    let sub_dir = current_dir.join("pre-configs");
    let mut options: Vec<String> = Vec::new();

    // Add option to remove service (will always be first)
    options.push("УДАЛИТЬ СЛУЖБУ С АВТОЗАПУСКА".to_string());

    // Add option for blockcheck (will always be second)
    options.push("ЗАПУСТИТЬ BLOCKCHECK (АВТО-ПОИСК)".to_string());

    // Collect and sort .bat files
    if let Ok(read_dir) = fs::read_dir(&sub_dir) {
        let mut bat_files: Vec<String> = read_dir
            .filter_map(|entry| {
                entry.ok().and_then(|e| {
                    let path = e.path();
                    if path.extension().and_then(|ext| ext.to_str()) == Some("bat") {
                        path.file_name().and_then(|n| n.to_str()).map(String::from)
                    } else {
                        None
                    }
                })
            })
            .collect();

        // Custom sorting for hierarchical file naming
        bat_files.sort_by(|a, b| {
            // Helper function to get parts of the filename
            fn split_filename(name: &str) -> (String, String, String, bool) {
                let without_ext = name.trim_end_matches(".bat");

                // Split into base name and parentheses part
                let (main_part, parentheses) = match without_ext.find('(') {
                    Some(idx) => (&without_ext[..idx - 1], &without_ext[idx..]),
                    None => (without_ext, ""),
                };

                // Split main part into components by underscore
                let parts: Vec<&str> = main_part.split('_').collect();
                let base = parts[0].to_string();

                // Get variant part (ALT, v2, etc)
                let variant = parts[1..].join("_");

                // Check if it's a provider variant
                let has_provider = !parentheses.is_empty();

                (base, variant, parentheses.to_string(), has_provider)
            }

            let (a_base, a_variant, a_provider, a_has_provider) = split_filename(a);
            let (b_base, b_variant, b_provider, b_has_provider) = split_filename(b);

            // First compare by base name
            match a_base.cmp(&b_base) {
                std::cmp::Ordering::Equal => {
                    // Same base name, compare variants
                    match a_variant.cmp(&b_variant) {
                        std::cmp::Ordering::Equal => {
                            // Same variant, non-provider version comes first
                            match (a_has_provider, b_has_provider) {
                                (false, true) => std::cmp::Ordering::Less,
                                (true, false) => std::cmp::Ordering::Greater,
                                // Both have or don't have providers, sort by provider name
                                _ => a_provider.cmp(&b_provider)
                            }
                        }
                        // Different variants
                        other => other
                    }
                }
                // Different base names
                other => other
            }
        });

        // Add sorted files to options
        options.extend(bat_files);
    }

    options
}

/// Renders the list of options to the terminal with scrolling support.
fn render_options(
    stdout: &mut impl Write,
    options: &[String],
    current_selection: usize,
    start_row: usize,
    scroll_offset: usize,
    max_visible_options: usize,
) -> crossterm::Result<()> {
    // Clear the options area
    execute!(
        stdout,
        cursor::MoveTo(0, start_row as u16),
        terminal::Clear(ClearType::FromCursorDown)
    )?;

    let mut current_row = start_row;

    // Display scroll indicator if there are options above
    if scroll_offset > 0 {
        execute!(
            stdout,
            cursor::MoveTo(0, current_row as u16),
            SetForegroundColor(Color::DarkGrey),
        )?;
        write!(stdout, "↑ Еще опции выше")?;
        execute!(stdout, ResetColor)?;
        current_row += 1;
    }

    // Calculate the visible range
    let end_index = std::cmp::min(scroll_offset + max_visible_options, options.len());
    let visible_range = scroll_offset..end_index;

    // Display visible options
    for (i, option) in options.iter().enumerate() {
        if visible_range.contains(&i) {
            execute!(stdout, cursor::MoveTo(0, current_row as u16))?;
            if i == current_selection {
                write!(stdout, "> {}", option)?;
            } else {
                write!(stdout, "  {}", option)?;
            }
            current_row += 1;
        }
    }

    // Display scroll indicator if there are options below
    if end_index < options.len() {
        execute!(
            stdout,
            cursor::MoveTo(0, current_row as u16),
            SetForegroundColor(Color::DarkGrey),
        )?;
        write!(stdout, "↓ Еще опции ниже")?;
        execute!(stdout, ResetColor)?;
    }

    stdout.flush()?;
    Ok(())
}

/// Handles the removal of the service.
fn handle_service_removal(
    stdout: &mut impl Write,
    mut message_row: usize,
) -> crossterm::Result<usize> {
    println!("Остановка и удаление службы 'discordfix_zapret'...");
    stdout.flush()?;

    // Stop the service
    message_row += 1;
    execute!(stdout, cursor::MoveTo(0, message_row as u16))?;
    match run_powershell_command(
        "Start-Process 'sc.exe' -ArgumentList 'stop discordfix_zapret' -Verb RunAs",
    ) {
        Ok(_) => println!("Служба 'discordfix_zapret' успешно остановлена."),
        Err(e) => println!("Ошибка при остановке службы: {}", e),
    }

    // Terminate the 'winws.exe' process
    message_row += 1;
    execute!(stdout, cursor::MoveTo(0, message_row as u16))?;
    println!("Завершение процесса 'winws.exe'...");
    stdout.flush()?;
    match run_powershell_command(
        "Start-Process 'powershell' -ArgumentList 'Stop-Process -Name \"winws\" -Force' -Verb RunAs",
    ) {
        Ok(_) => println!("Процесс 'winws.exe' успешно завершён."),
        Err(e) => println!("Ошибка при завершении процесса 'winws.exe': {}", e),
    }

    // Delete the service
    message_row += 1;
    execute!(stdout, cursor::MoveTo(0, message_row as u16))?;
    match run_powershell_command(
        "Start-Process 'sc.exe' -ArgumentList 'delete discordfix_zapret' -Verb RunAs",
    ) {
        Ok(_) => println!("Служба 'discordfix_zapret' успешно удалена."),
        Err(e) => println!("Ошибка при удалении службы: {}", e),
    }

    stdout.flush()?;
    Ok(message_row)
}

/// Handles the installation of the selected .bat file as a service.
fn handle_service_installation(
    stdout: &mut impl Write,
    selected_file: &str,
    mut message_row: usize,
) -> crossterm::Result<usize> {
    // Stop and remove any existing service
    message_row = handle_service_removal(stdout, message_row)?;

    // Install the selected .bat file as a service
    message_row += 1;
    execute!(stdout, cursor::MoveTo(0, message_row as u16))?;
    println!("\nУстановка файла как службы: {}", selected_file);
    stdout.flush()?;

    let current_dir = env::current_dir().expect("Не удалось получить текущую директорию");
    let sub_dir = current_dir.join("pre-configs");
    let bat_file_path = sub_dir.join(selected_file);
    let service_name = "discordfix_zapret";

    // Create the service
    message_row += 2;
    execute!(stdout, cursor::MoveTo(0, message_row as u16))?;
    let create_command = format!(
        "Start-Process 'sc.exe' -ArgumentList 'create {} binPath= \"cmd.exe /c \"\"{}\"\"\" start= auto' -Verb RunAs",
        service_name,
        bat_file_path.display()
    );
    match run_powershell_command(&create_command) {
        Ok(_) => println!("Служба успешно установлена."),
        Err(e) => println!("Ошибка при установке службы: {}", e),
    }

    // Start the service
    message_row += 1;
    execute!(stdout, cursor::MoveTo(0, message_row as u16))?;
    let start_command = format!(
        "Start-Process 'sc.exe' -ArgumentList 'start {}' -Verb RunAs",
        service_name
    );
    match run_powershell_command(&start_command) {
        Ok(_) => println!("Служба успешно запущена."),
        Err(e) => println!("Ошибка при запуске службы: {}", e),
    }

    stdout.flush()?;
    Ok(message_row)
}

/// Runs a PowerShell command and captures its output.
fn run_powershell_command(command: &str) -> Result<(), String> {
    let output = Command::new("powershell")
        .args(&["-Command", command])
        .output()
        .map_err(|e| format!("Не удалось выполнить команду: {}", e))?;

    if output.status.success() {
        Ok(())
    } else {
        Err(String::from_utf8_lossy(&output.stderr).into_owned())
    }
}

/// Cleans up the terminal by disabling raw mode and restoring cursor visibility.
fn cleanup_terminal() -> crossterm::Result<()> {
    terminal::disable_raw_mode()?;
    execute!(stdout(), terminal::LeaveAlternateScreen, cursor::Show)?;
    Ok(())
}
