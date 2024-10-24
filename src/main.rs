use std::{
    env,
    fs,
    io::{self, stdout, Write},
    process::Command,
    time::{Duration, Instant},
};

use crossterm::{
    cursor,
    event::{self, Event, KeyCode},
    execute,
    style::{Color, Print, ResetColor, SetForegroundColor},
    terminal::{self, Clear, ClearType},
};

fn main() -> io::Result<()> {
    let original_hook = std::panic::take_hook();
    std::panic::set_hook(Box::new(move |panic_info| {
        let _ = cleanup_terminal();
        original_hook(panic_info);
    }));

    let mut stdout = stdout();

    // Ensure clean terminal state
    terminal::enable_raw_mode()?;
    execute!(
        stdout,
        terminal::EnterAlternateScreen,
        cursor::Hide,
        Clear(ClearType::All),
        cursor::MoveTo(0, 0)
    )?;

    // Get terminal size early
    let (_, term_height) = terminal::size()?;

    // Print welcome messages and get the starting line
    let current_line = print_welcome_message();

    // Get the list of .bat files and add options
    let options = get_options();
    if options.is_empty() {
        println!("Не найдено ни одного .bat файла в текущей директории.");
        cleanup_terminal()?;
        return Ok(());
    }

    let mut current_selection = 0;
    let start_row = current_line;

    // Calculate maximum visible options, reserving space for UI elements
    // Calculate maximum visible options, limiting to 10
    let max_visible_options = std::cmp::min(
        15,
        (term_height - start_row as u16 - 3) as usize
    );

    let mut scroll_offset = 0;

    // Initial render
    render_options(
        &mut stdout,
        &options,
        current_selection,
        start_row,
        scroll_offset,
        max_visible_options,
    )?;

    // Variable to keep track of where to print messages
    let mut message_row = start_row + std::cmp::min(options.len(), max_visible_options) + 1;

    // Variable to track the time of the last key event
    let mut last_event_time = Instant::now();

    // Main event loop
    loop {
        if event::poll(Duration::from_millis(100))? {
            if let Event::Key(event) = event::read()? {
                let now = Instant::now();
                if now.duration_since(last_event_time) > Duration::from_millis(150) {
                    match event.code {
                        KeyCode::Up => {
                            if current_selection > 0 {
                                current_selection -= 1;
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
                                let visible_options = max_visible_options.saturating_sub(2);
                                if current_selection >= scroll_offset + visible_options {
                                    scroll_offset = current_selection.saturating_sub(visible_options - 1);
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
                            let visible_options = max_visible_options.saturating_sub(2);
                            current_selection = current_selection.saturating_sub(visible_options);
                            scroll_offset = scroll_offset.saturating_sub(visible_options);
                            render_options(
                                &mut stdout,
                                &options,
                                current_selection,
                                start_row,
                                scroll_offset,
                                max_visible_options,
                            )?;
                        }
                        KeyCode::PageDown => {
                            let visible_options = max_visible_options.saturating_sub(2);
                            current_selection = (current_selection + visible_options).min(options.len() - 1);
                            scroll_offset = (scroll_offset + visible_options).min(
                                options.len().saturating_sub(visible_options)
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
                        KeyCode::Enter => {
                            // Clear messages below the options
                            execute!(
                                stdout,
                                cursor::MoveTo(0, message_row as u16),
                                Clear(ClearType::FromCursorDown)
                            )?;
                            stdout.flush()?;

                            if options[current_selection] == "УДАЛИТЬ СЛУЖБУ С АВТОЗАПУСКА" {
                                message_row = handle_service_removal(&mut stdout, message_row)?;
                            } else if options[current_selection] == "ЗАПУСТИТЬ BLOCKCHECK (АВТО-ПОДБОР-ПАРАМЕТРОВ-БАТНИКА)" {
                                message_row += 1;
                                execute!(stdout, cursor::MoveTo(0, message_row as u16))?;
                                match run_powershell_command("Start-Process 'blockcheck.cmd'") {
                                    Ok(_) => {
                                        println!("Blockcheck успешно запущен.");
                                        std::process::exit(0);
                                    },
                                    Err(e) => println!("Ошибка при запуске Blockcheck: {}", e),
                                }
                            } else {
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

    execute!(
        stdout,
        cursor::MoveTo(0, message_row as u16 + 2),
        Clear(ClearType::FromCursorDown)
    )?;

    // Clean up terminal before final messages
    cleanup_terminal()?;

    println!("\nГотово!\nВы можете закрыть это окно.");

    // Wait for Enter key in normal terminal mode
    let mut input = String::new();
    io::stdin().read_line(&mut input)?;

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
        'УДАЛИТЬ СЛУЖБУ С АВТОЗАПУСКА' или 'ЗАПУСТИТЬ BLOCKCHECK (АВТО-ПОДБОР-ПАРАМЕТРОВ-БАТНИКА)'.\n"
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
    options.push("ЗАПУСТИТЬ BLOCKCHECK (АВТО-ПОДБОР-ПАРАМЕТРОВ-БАТНИКА)".to_string());

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
) -> io::Result<()> {
    // Limit visible options to 10 (8 options + 2 scroll indicators)
    let max_display_options = 15;
    let visible_options = std::cmp::min(max_display_options - 2, max_visible_options.saturating_sub(2));
    let total_options = options.len();

    // Adjust scroll_offset to keep selection visible within the 10-item window
    let adjusted_scroll_offset = if current_selection >= scroll_offset + visible_options {
        current_selection.saturating_sub(visible_options - 1)
    } else if current_selection < scroll_offset {
        current_selection
    } else {
        scroll_offset
    };

    // Calculate the visible range
    let end_index = (adjusted_scroll_offset + visible_options).min(total_options);
    let visible_range = adjusted_scroll_offset..end_index;

    // Constants for formatting
    const MARKER: &str = "►";
    const EMPTY_MARKER: &str = " ";
    const SPACING: &str = " ";

    // Clear the entire options area first
    execute!(
        stdout,
        cursor::MoveTo(0, start_row as u16),
        Clear(ClearType::FromCursorDown)
    )?;

    let mut current_row = start_row;

    // Up scroll indicator (if we're not at the top)
    if adjusted_scroll_offset > 0 {
        execute!(
            stdout,
            cursor::MoveTo(0, current_row as u16),
            Clear(ClearType::CurrentLine),
            SetForegroundColor(Color::DarkGrey),
            Print("↑ Еще опции выше"),
            ResetColor
        )?;
        current_row += 1;
    }

    // Display visible options (limited to our window size)
    let displayed_count = 0;
    for (index, option) in options.iter().enumerate() {
        if visible_range.contains(&index) && displayed_count < visible_options {
            // Position cursor at start of line
            execute!(
                stdout,
                cursor::MoveTo(0, current_row as u16),
                Clear(ClearType::CurrentLine)
            )?;

            if index == current_selection {
                // Selected item - print marker and highlighted text
                execute!(
                    stdout,
                    SetForegroundColor(Color::Cyan),
                    Print(MARKER),
                    Print(SPACING),
                    Print(option),
                    ResetColor
                )?;
            } else {
                // Unselected item - print empty marker and normal text
                execute!(
                    stdout,
                    Print(EMPTY_MARKER),
                    Print(SPACING),
                    Print(option)
                )?;
            }

            stdout.flush()?;
            current_row += 1;
        }
    }

    // Down scroll indicator (if we're not at the bottom)
    if end_index < total_options {
        execute!(
            stdout,
            cursor::MoveTo(0, current_row as u16),
            Clear(ClearType::CurrentLine),
            SetForegroundColor(Color::DarkGrey),
            Print("↓ Еще опции ниже"),
            ResetColor
        )?;
    }

    // Final flush to ensure everything is rendered
    stdout.flush()?;
    Ok(())
}

/// Handles the removal of the service.
fn handle_service_removal(
    stdout: &mut impl Write,
    mut message_row: usize,
) -> io::Result<usize> {
    // Initial message
    execute!(
        stdout,
        cursor::MoveTo(0, message_row as u16),
        Clear(ClearType::FromCursorDown),
        Print("=== Удаление существующей службы ===")
    )?;
    message_row += 1;
    stdout.flush()?;

    // Stop service message
    execute!(
        stdout,
        cursor::MoveTo(0, message_row as u16),
        Clear(ClearType::CurrentLine),
        Print("► Остановка службы 'discordfix_zapret'...")
    )?;
    stdout.flush()?;

    // Stop the service
    match run_powershell_command(
        "Start-Process 'sc.exe' -ArgumentList 'stop discordfix_zapret' -Verb RunAs",
    ) {
        Ok(_) => {
            message_row += 1;
            execute!(
                stdout,
                cursor::MoveTo(0, message_row as u16),
                Clear(ClearType::CurrentLine),
                Print("✓ Служба 'discordfix_zapret' успешно остановлена.")
            )?;
        },
        Err(e) => {
            message_row += 1;
            execute!(
                stdout,
                cursor::MoveTo(0, message_row as u16),
                Clear(ClearType::CurrentLine),
                Print(format!("⚠ Ошибка при остановке службы: {}", e))
            )?;
        }
    }
    stdout.flush()?;

    // Terminate process message
    message_row += 1;
    execute!(
        stdout,
        cursor::MoveTo(0, message_row as u16),
        Clear(ClearType::CurrentLine),
        Print("► Завершение процесса 'winws.exe'...")
    )?;
    stdout.flush()?;

    // Terminate process
    match run_powershell_command(
        "Start-Process 'powershell' -ArgumentList 'Stop-Process -Name \"winws\" -Force' -Verb RunAs",
    ) {
        Ok(_) => {
            message_row += 1;
            execute!(
                stdout,
                cursor::MoveTo(0, message_row as u16),
                Clear(ClearType::CurrentLine),
                Print("✓ Процесс 'winws.exe' успешно завершён.")
            )?;
        },
        Err(e) => {
            message_row += 1;
            execute!(
                stdout,
                cursor::MoveTo(0, message_row as u16),
                Clear(ClearType::CurrentLine),
                Print(format!("⚠ Ошибка при завершении процесса 'winws.exe': {}", e))
            )?;
        }
    }
    stdout.flush()?;

    // Delete service message
    message_row += 1;
    execute!(
        stdout,
        cursor::MoveTo(0, message_row as u16),
        Clear(ClearType::CurrentLine),
        Print("► Удаление службы 'discordfix_zapret'...")
    )?;
    stdout.flush()?;

    // Delete the service
    match run_powershell_command(
        "Start-Process 'sc.exe' -ArgumentList 'delete discordfix_zapret' -Verb RunAs",
    ) {
        Ok(_) => {
            message_row += 1;
            execute!(
                stdout,
                cursor::MoveTo(0, message_row as u16),
                Clear(ClearType::CurrentLine),
                Print("✓ Служба 'discordfix_zapret' успешно удалена.")
            )?;
        },
        Err(e) => {
            message_row += 1;
            execute!(
                stdout,
                cursor::MoveTo(0, message_row as u16),
                Clear(ClearType::CurrentLine),
                Print(format!("⚠ Ошибка при удалении службы: {}", e))
            )?;
        }
    }
    stdout.flush()?;

    message_row += 1;
    Ok(message_row)
}

/// Handles the installation of the selected .bat file as a service.
fn handle_service_installation(
    stdout: &mut impl Write,
    selected_file: &str,
    mut message_row: usize,
) -> io::Result<usize> {
    // First remove any existing service
    message_row = handle_service_removal(stdout, message_row)?;

    // Add spacing after removal messages
    message_row += 1;
    execute!(
        stdout,
        cursor::MoveTo(0, message_row as u16),
        Clear(ClearType::CurrentLine),
        Print("=== Установка новой службы ===")
    )?;
    message_row += 1;

    execute!(
        stdout,
        cursor::MoveTo(0, message_row as u16),
        Clear(ClearType::CurrentLine),
        Print(format!("► Установка файла как службы: {}", selected_file))
    )?;
    stdout.flush()?;

    let current_dir = env::current_dir().expect("Не удалось получить текущую директорию");
    let sub_dir = current_dir.join("pre-configs");
    let bat_file_path = sub_dir.join(selected_file);
    let service_name = "discordfix_zapret";

    // Create service
    message_row += 1;
    execute!(
        stdout,
        cursor::MoveTo(0, message_row as u16),
        Clear(ClearType::CurrentLine),
        Print("► Создание службы...")
    )?;
    stdout.flush()?;

    let create_command = format!(
        "Start-Process 'sc.exe' -ArgumentList 'create {} binPath= \"cmd.exe /c \"\"{}\"\"\" start= auto' -Verb RunAs",
        service_name,
        bat_file_path.display()
    );

    match run_powershell_command(&create_command) {
        Ok(_) => {
            message_row += 1;
            execute!(
                stdout,
                cursor::MoveTo(0, message_row as u16),
                Clear(ClearType::CurrentLine),
                Print("✓ Служба успешно установлена.")
            )?;
        },
        Err(e) => {
            message_row += 1;
            execute!(
                stdout,
                cursor::MoveTo(0, message_row as u16),
                Clear(ClearType::CurrentLine),
                Print(format!("⚠ Ошибка при установке службы: {}", e))
            )?;
        }
    }

    // Start service
    message_row += 1;
    execute!(
        stdout,
        cursor::MoveTo(0, message_row as u16),
        Clear(ClearType::CurrentLine),
        Print("► Запуск службы...")
    )?;
    stdout.flush()?;

    let start_command = format!(
        "Start-Process 'sc.exe' -ArgumentList 'start {}' -Verb RunAs",
        service_name
    );

    match run_powershell_command(&start_command) {
        Ok(_) => {
            message_row += 1;
            execute!(
                stdout,
                cursor::MoveTo(0, message_row as u16),
                Clear(ClearType::CurrentLine),
                Print("✓ Служба успешно запущена.")
            )?;
        },
        Err(e) => {
            message_row += 1;
            execute!(
                stdout,
                cursor::MoveTo(0, message_row as u16),
                Clear(ClearType::CurrentLine),
                Print(format!("⚠ Ошибка при запуске службы: {}", e))
            )?;
        }
    }

    message_row += 1;
    stdout.flush()?;
    Ok(message_row)
}

/// Runs a PowerShell command and captures its output.
fn run_powershell_command(command: &str) -> io::Result<()> {
    let output = Command::new("powershell")
        .args(&["-Command", command])
        .output()
        .map_err(|e| io::Error::new(
            io::ErrorKind::Other,
            format!("Не удалось выполнить команду: {}", e)
        ))?;

    if output.status.success() {
        Ok(())
    } else {
        // Convert stderr to string, handle invalid UTF-8 gracefully
        let error_message = String::from_utf8_lossy(&output.stderr).into_owned();

        // If error message is empty, try to get anything from stdout
        let error_message = if error_message.is_empty() {
            String::from_utf8_lossy(&output.stdout).into_owned()
        } else {
            error_message
        };

        // If both stderr and stdout are empty, provide a generic error message
        let error_message = if error_message.is_empty() {
            "Неизвестная ошибка при выполнении команды PowerShell".to_string()
        } else {
            error_message
        };

        Err(io::Error::new(io::ErrorKind::Other, error_message))
    }
}

// Enhanced cleanup function
fn cleanup_terminal() -> io::Result<()> {
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

// Add this helper function to ensure proper terminal cleanup on exit
fn setup_terminal_cleanup() {
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

fn clear_screen(stdout: &mut impl Write) -> io::Result<()> {
    execute!(
        stdout,
        Clear(ClearType::All),
        cursor::MoveTo(0, 0)
    )?;
    stdout.flush()?;
    Ok(())
}
