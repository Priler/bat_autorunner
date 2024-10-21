use crossterm::{
    cursor,
    event::{self, Event, KeyCode},
    execute,
    terminal::{self, ClearType},
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

    // Get the list of .bat files and add the option to remove the service
    let options = get_options();
    if options.is_empty() {
        println!("Не найдено ни одного .bat файла в текущей директории.");
        cleanup_terminal()?;
        return Ok(());
    }

    let mut current_selection = 0;
    let start_row = current_line; // Start rendering options from this line

    // Render the initial list of options
    render_options(&mut stdout, &options, current_selection, start_row)?;

    // Variable to keep track of where to print messages
    let mut message_row = start_row + options.len();

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
                                render_options(
                                    &mut stdout,
                                    &options,
                                    current_selection,
                                    start_row,
                                )?;
                            }
                        }
                        KeyCode::Down => {
                            if current_selection < options.len() - 1 {
                                current_selection += 1;
                                render_options(
                                    &mut stdout,
                                    &options,
                                    current_selection,
                                    start_row,
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
                                // Remove the service
                                message_row =
                                    handle_service_removal(&mut stdout, message_row)?;
                            } else {
                                // Install the selected .bat file as a service
                                let selected_file = &options[current_selection];
                                message_row = handle_service_installation(
                                    &mut stdout,
                                    selected_file,
                                    message_row,
                                )?;
                            }

                            // Break after processing the selection
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

    // Cleanup terminal settings and exit
    // cleanup_terminal()?;
    println!("\nГотово! Вы можете закрыть это окно.");
    println!("Нажмите Enter для выхода.");

    // Wait for user input before closing
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
    println!("Версия: 0.1.1");
    current_line += 1;
    println!("===");
    current_line += 1;
    println!(
        "Используя СТРЕЛКИ на клавиатуре, выберите .bat файл из списка \
        для установки службы 'discordfix_zapret' или выберите \
        'УДАЛИТЬ СЛУЖБУ С АВТОЗАПУСКА'.\n"
    );
    current_line += 2; // Account for the extra newline
    println!("Для выбора нажмите ENTER.");
    current_line += 1;
    current_line
}

/// Retrieves a list of .bat files in the current directory and adds the remove service option.
fn get_options() -> Vec<String> {
    let current_dir = env::current_dir().expect("Не удалось получить текущую директорию");
    let mut options: Vec<String> = Vec::new();

    match fs::read_dir(&current_dir) {
        Ok(read_dir) => {
            for entry in read_dir {
                if let Ok(entry) = entry {
                    let path = entry.path();
                    if path.extension().and_then(|ext| ext.to_str()) == Some("bat") {
                        if let Some(name) = path.file_name().and_then(|n| n.to_str()) {
                            options.push(name.to_string());
                        }
                    }
                }
            }
        }
        Err(_) => {
            // Handle the error (e.g., log it or ignore it)
            // In this case, we simply return an empty options vector
        }
    }

    // Add option to remove service
    options.push("УДАЛИТЬ СЛУЖБУ С АВТОЗАПУСКА".to_string());

    options
}

/// Renders the list of options to the terminal.
fn render_options(
    stdout: &mut impl Write,
    options: &[String],
    current_selection: usize,
    start_row: usize,
) -> crossterm::Result<()> {
    execute!(
        stdout,
        cursor::MoveTo(0, start_row as u16),
        terminal::Clear(ClearType::FromCursorDown)
    )?;
    for (i, option) in options.iter().enumerate() {
        if i == current_selection {
            println!("> {}", option);
        } else {
            println!("  {}", option);
        }
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
    let bat_file_path = current_dir.join(selected_file);
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
