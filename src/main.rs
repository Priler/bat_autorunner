use crossterm::{
    cursor,
    event::{self, Event, KeyCode},
    execute,
    terminal::{self, ClearType},
};
use std::fs;
use std::io::{stdout, Write};
use std::process::Command;
use std::time::{Duration, Instant};
use std::thread::sleep;
use std::env;
use std::path::PathBuf;
use std::io;

fn main() -> crossterm::Result<()> {
    let mut stdout = stdout();
    execute!(stdout, terminal::EnterAlternateScreen)?;
    terminal::enable_raw_mode()?;

    // Hello message
    println!("Добро пожаловать!");
    println!("Эта программа позволяет установить .bat файл как службу с автозапуском.");
    println!("Автор программы: Хауди (Абрахам) https://github.com/Priler");
    println!("===");
    println!("Используя СТРЕЛКИ на клавиатуре, выберите .bat файл из списка для установки службы 'discordfix_zapret' или выберите 'УДАЛИТЬ СЛУЖБУ С АВТОЗАПУСКА'.\n");
    println!("Для выбора нажмите ENTER.");

    // Get current dir
    let current_dir = env::current_dir().expect("Не удалось получить текущую директорию");

    // Locate .bat files
    let mut options: Vec<String> = fs::read_dir(current_dir.clone())
        .unwrap()
        .filter_map(|entry| {
            let entry = entry.unwrap();
            let path = entry.path();
            if path.extension().map_or(false, |ext| ext == "bat") {
                Some(path.file_name().unwrap().to_string_lossy().into_owned())
            } else {
                None
            }
        })
        .collect();

    // Add option to remove service
    options.push("УДАЛИТЬ СЛУЖБУ С АВТОЗАПУСКА".to_string());

    if options.is_empty() {
        println!("Не найдено ни одного .bat файла в текущей директории.");
        terminal::disable_raw_mode()?;
        execute!(stdout, terminal::LeaveAlternateScreen)?;
        return Ok(());
    }

    let mut current_selection = 0;
    let mut last_selection = 0;
    let mut last_event_time = Instant::now();
    let start_row = 10;

    // Print all options
    for (i, option) in options.iter().enumerate() {
        execute!(stdout, cursor::MoveTo(0, (start_row + i) as u16))?;
        if i == current_selection {
            println!("> {}", option);
        } else {
            println!("  {}", option);
        }
    }

    let mut message_row = (start_row + options.len()) as u16;

    loop {
        if event::poll(Duration::from_millis(100))? {
            if let Event::Key(event) = event::read()? {
                let now = Instant::now();
                if now.duration_since(last_event_time) > Duration::from_millis(150) {
                    match event.code {
                        KeyCode::Up => {
                            if current_selection > 0 {
                                current_selection -= 1;
                            }
                        }
                        KeyCode::Down => {
                            if current_selection < options.len() - 1 {
                                current_selection += 1;
                            }
                        }
                        KeyCode::Enter => {
                            execute!(stdout, cursor::MoveTo(0, message_row), terminal::Clear(ClearType::FromCursorDown))?;

                            // Remove all service option
                            if options[current_selection] == "УДАЛИТЬ СЛУЖБУ С АВТОЗАПУСКА" {
                                println!("Остановка и удаление службы 'discordfix_zapret'...");

                                // Остановка службы через PowerShell с Verb RunAs
                                let stop_output = Command::new("powershell")
                                    .arg("-Command")
                                    .arg("Start-Process 'sc.exe' -ArgumentList 'stop discordfix_zapret' -Verb RunAs")
                                    .output()
                                    .expect("Не удалось выполнить команду остановки службы");

                                // Перемещаем курсор вниз на новую строку для следующего сообщения
                                message_row += 1;
                                execute!(stdout, cursor::MoveTo(0, message_row))?;
                                if stop_output.status.success() {
                                    println!("Служба 'discordfix_zapret' успешно остановлена.");
                                } else {
                                    println!("Ошибка при остановке службы: {:?}", String::from_utf8_lossy(&stop_output.stderr));
                                }

                                // Завершаем процесс с именем "winws.exe" через PowerShell с правами администратора
                                println!("Завершение процесса 'winws.exe'...");
                                let kill_output = Command::new("powershell")
                                    .arg("-Command")
                                    .arg("Start-Process 'powershell' -ArgumentList 'Stop-Process -Name \"winws\" -Force' -Verb RunAs")
                                    .output()
                                    .expect("Не удалось завершить процесс 'winws.exe'");

                                // Перемещаем курсор на новую строку для следующего сообщения
                                message_row += 1;
                                execute!(stdout, cursor::MoveTo(0, message_row))?;
                                if kill_output.status.success() {
                                    println!("Процесс 'winws.exe' успешно завершён.");
                                } else {
                                    println!("Ошибка при завершении процесса 'winws.exe': {:?}", String::from_utf8_lossy(&kill_output.stderr));
                                }

                                // Удаление службы через PowerShell с Verb RunAs
                                let delete_output = Command::new("powershell")
                                    .arg("-Command")
                                    .arg("Start-Process 'sc.exe' -ArgumentList 'delete discordfix_zapret' -Verb RunAs")
                                    .output()
                                    .expect("Не удалось выполнить команду удаления службы");

                                // Перемещаем курсор на новую строку для следующего сообщения
                                message_row += 1;
                                execute!(stdout, cursor::MoveTo(0, message_row))?;
                                if delete_output.status.success() {
                                    println!("Служба 'discordfix_zapret' успешно удалена.");
                                } else {
                                    println!("Ошибка при удалении службы: {:?}", String::from_utf8_lossy(&delete_output.stderr));
                                }

                                // Гарантируем, что сообщение отобразится
                                stdout.flush()?;

                                break;
                            } else {
                                println!("Остановка и удаление службы 'discordfix_zapret'...");

                                // Stop service with PowerShell + Verb RunAs
                                let stop_output = Command::new("powershell")
                                    .arg("-Command")
                                    .arg("Start-Process 'sc.exe' -ArgumentList 'stop discordfix_zapret' -Verb RunAs")
                                    .output()
                                    .expect("Не удалось выполнить команду остановки службы");

                                // Move cursor
                                message_row += 1;
                                execute!(stdout, cursor::MoveTo(0, message_row))?;
                                if stop_output.status.success() {
                                    println!("Служба 'discordfix_zapret' успешно остановлена.");
                                } else {
                                    println!("Ошибка при остановке службы: {:?}", String::from_utf8_lossy(&stop_output.stderr));
                                }

                                // Stop process "winws.exe" with PowerShell + Verb RunAs
                                println!("Завершение процесса 'winws.exe'...");
                                let kill_output = Command::new("powershell")
                                    .arg("-Command")
                                    .arg("Start-Process 'powershell' -ArgumentList 'Stop-Process -Name \"winws\" -Force' -Verb RunAs")
                                    .output()
                                    .expect("Не удалось завершить процесс 'winws.exe'");

                                // Move cursor
                                message_row += 1;
                                execute!(stdout, cursor::MoveTo(0, message_row))?;
                                if kill_output.status.success() {
                                    println!("Процесс 'winws.exe' успешно завершён.");
                                } else {
                                    println!("Ошибка при завершении процесса 'winws.exe': {:?}", String::from_utf8_lossy(&kill_output.stderr));
                                }

                                // Remove service with PowerShell + Verb RunAs
                                let delete_output = Command::new("powershell")
                                    .arg("-Command")
                                    .arg("Start-Process 'sc.exe' -ArgumentList 'delete discordfix_zapret' -Verb RunAs")
                                    .output()
                                    .expect("Не удалось выполнить команду удаления службы");

                                // Move cursor
                                message_row += 1;
                                execute!(stdout, cursor::MoveTo(0, message_row))?;
                                if delete_output.status.success() {
                                    println!("Служба 'discordfix_zapret' успешно удалена.");
                                } else {
                                    println!("Ошибка при удалении службы: {:?}", String::from_utf8_lossy(&delete_output.stderr));
                                }

                                // Install selected .bat file as service called "discordfix_zapret"
                                let selected_file = options[current_selection].clone();
                                println!("\nУстановка файла как службы: {}", selected_file);

                                // Move cursor
                                message_row += 2;
                                execute!(stdout, cursor::MoveTo(0, message_row))?;

                                // Full path to selected .bat file
                                let bat_file_path = current_dir.join(&selected_file);

                                // Service name is "discordfix_zapret"
                                let service_name = "discordfix_zapret";

                                // Install service with PowerShell + Verb RunAs
                                let create_output = Command::new("powershell")
                                    .arg("-Command")
                                    .arg(format!(
                                        "Start-Process 'sc.exe' -ArgumentList 'create {} binPath= \"cmd.exe /c {}\" start= auto' -Verb RunAs",
                                        service_name,
                                        bat_file_path.display()
                                    ))
                                    .output()
                                    .expect("Не удалось выполнить команду создания службы");

                                // Move cursor
                                message_row += 1;
                                execute!(stdout, cursor::MoveTo(0, message_row))?;
                                if create_output.status.success() {
                                    println!("Служба успешно установлена.");
                                } else {
                                    println!("Ошибка при установке службы: {:?}", String::from_utf8_lossy(&create_output.stderr));
                                }

                                // Start service with PowerShell + Verb RunAs
                                let start_output = Command::new("powershell")
                                    .arg("-Command")
                                    .arg(format!(
                                        "Start-Process 'sc.exe' -ArgumentList 'start {}' -Verb RunAs",
                                        service_name
                                    ))
                                    .output()
                                    .expect("Не удалось выполнить команду запуска службы");

                                // Move cursor
                                message_row += 1;
                                execute!(stdout, cursor::MoveTo(0, message_row))?;
                                if start_output.status.success() {
                                    println!("Служба успешно запущена.");
                                } else {
                                    println!("Ошибка при запуске службы: {:?}", String::from_utf8_lossy(&start_output.stderr));
                                }

                                // Flush screen
                                stdout.flush()?;
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

        // Обновляем только изменившийся элемент
        if current_selection != last_selection {
            // Убираем старое выделение
            execute!(stdout, cursor::MoveTo(0, (start_row + last_selection) as u16))?;
            print!("  {}", options[last_selection]);
            stdout.flush()?;

            // Показываем новое выделение
            execute!(stdout, cursor::MoveTo(0, (start_row + current_selection) as u16))?;
            print!("> {}", options[current_selection]);
            stdout.flush()?;

            last_selection = current_selection;
        }

        sleep(Duration::from_millis(50));
    }

    // Отключаем raw режим консоли
    terminal::disable_raw_mode()?;

    // Выводим завершающее сообщение на следующей строке под предыдущими сообщениями
    execute!(stdout, cursor::MoveTo(0, message_row))?;
    println!("\nГотово! Вы можете закрыть это окно.");
    println!("Нажмите Enter для выхода.");

    // Пауза перед закрытием
    let mut _input = String::new();
    io::stdin().read_line(&mut _input).expect("Ошибка при ожидании ввода.");

    Ok(())
}
