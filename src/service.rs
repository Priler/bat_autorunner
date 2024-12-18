use std::io::{self, Write};
use crossterm::{
    cursor,
    execute,
    style::{Print, SetForegroundColor, Color, ResetColor},
};
use crate::utils::run_powershell_command_with_output;
use regex::Regex;

pub struct ServiceManager {
    service_name: String,
}

impl ServiceManager {
    pub fn new(service_name: &str) -> Self {
        Self {
            service_name: service_name.to_string(),
        }
    }

    pub fn remove_service(&self, stdout: &mut impl Write, mut message_row: usize) -> io::Result<usize> {
        // Section header
        execute!(
            stdout,
            cursor::MoveTo(0, message_row as u16),
            Print("\n=== Удаление существующей службы ===\n\n")
        )?;
        message_row += 3; // Account for the header and two newlines
        stdout.flush()?;

        // Process each operation
        message_row = self.stop_service(stdout, message_row)?;
        message_row = self.terminate_process(stdout, message_row)?;
        message_row = self.delete_service(stdout, message_row)?;

        Ok(message_row)
    }

    pub fn install_service(
        &self,
        stdout: &mut impl Write,
        bat_file_path: &str,
        mut message_row: usize
    ) -> io::Result<usize> {
        // First remove existing service
        message_row = self.remove_service(stdout, message_row)?;

        // Add spacing and section header for installation
        execute!(
            stdout,
            cursor::MoveTo(0, message_row as u16),
            Print("\n=== Установка новой службы ===\n\n")
        )?;
        message_row += 3;

        execute!(
            stdout,
            cursor::MoveTo(0, message_row as u16),
            Print(format!("► Установка файла как службы: {}\n", bat_file_path))
        )?;
        message_row += 1;
        stdout.flush()?;

        // Create and start service
        message_row = self.create_service(stdout, bat_file_path, message_row)?;
        message_row = self.start_service(stdout, message_row)?;

        Ok(message_row)
    }

    fn stop_service(&self, stdout: &mut impl Write, mut message_row: usize) -> io::Result<usize> {
        execute!(
            stdout,
            cursor::MoveTo(0, message_row as u16),
            Print(format!("► Остановка службы '{}'...\n", self.service_name))
        )?;
        message_row += 1;
        stdout.flush()?;

        let command = format!(
            "Start-Process 'sc.exe' -ArgumentList 'stop {}' -Verb RunAs",
            self.service_name
        );

        match run_powershell_command_with_output(&command) {
            Ok(_) => {
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Green),
                    Print(format!("✓ Служба '{}' успешно остановлена.\n", self.service_name)),
                    ResetColor
                )?;
                message_row += 2; // Add extra line for spacing
            },
            Err(e) => {
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Red),
                    Print(format!("⚠ Ошибка при остановке службы: {}\n", e)),
                    ResetColor
                )?;
                message_row += 2; // Add extra line for spacing
            }
        }

        Ok(message_row)
    }

    fn terminate_process(&self, stdout: &mut impl Write, mut message_row: usize) -> io::Result<usize> {
        execute!(
            stdout,
            cursor::MoveTo(0, message_row as u16),
            Print("► Завершение процесса 'winws.exe'...\n")
        )?;
        message_row += 1;
        stdout.flush()?;

        let command = "Start-Process 'powershell' -ArgumentList 'Stop-Process -Name \"winws\" -Force' -Verb RunAs";

        match run_powershell_command_with_output(command) {
            Ok(_) => {
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Green),
                    Print("✓ Процесс 'winws.exe' успешно завершён.\n"),
                    ResetColor
                )?;
                message_row += 2; // Add extra line for spacing
            },
            Err(e) => {
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Red),
                    Print(format!("⚠ Ошибка при завершении процесса 'winws.exe': {}\n", e)),
                    ResetColor
                )?;
                message_row += 2; // Add extra line for spacing
            }
        }

        Ok(message_row)
    }

    fn delete_service(&self, stdout: &mut impl Write, mut message_row: usize) -> io::Result<usize> {
        execute!(
            stdout,
            cursor::MoveTo(0, message_row as u16),
            Print(format!("► Удаление службы '{}'...\n", self.service_name))
        )?;
        message_row += 1;
        stdout.flush()?;

        let command = format!(
            "Start-Process 'sc.exe' -ArgumentList 'delete {}' -Verb RunAs",
            self.service_name
        );

        match run_powershell_command_with_output(&command) {
            Ok(_) => {
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Green),
                    Print(format!("✓ Служба '{}' успешно удалена.\n", self.service_name)),
                    ResetColor
                )?;
                message_row += 2; // Add extra line for spacing
            },
            Err(e) => {
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Red),
                    Print(format!("⚠ Ошибка при удалении службы: {}\n", e)),
                    ResetColor
                )?;
                message_row += 2; // Add extra line for spacing
            }
        }

        Ok(message_row)
    }

    fn create_service(&self, stdout: &mut impl Write, bat_file_path: &str, mut message_row: usize) -> io::Result<usize> {
        execute!(
            stdout,
            cursor::MoveTo(0, message_row as u16),
            Print("► Создание службы...\n")
        )?;
        message_row += 1;
        stdout.flush()?;

        // First, check if service exists
        let check_command = format!(
            "$service = Get-Service -Name '{}' -ErrorAction SilentlyContinue; if ($service) {{ Write-Output 'exists' }}",
            self.service_name
        );

        if let Ok(output) = run_powershell_command_with_output(&check_command) {
            if output.contains("exists") {
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Yellow),
                    Print("⚠ Служба уже существует. Попытка удаления...\n"),
                    ResetColor
                )?;
                message_row += 1;

                // Try to remove existing service
                let _ = self.remove_service(stdout, message_row)?;
            }
        }

        // Create service with proper path escaping and validation
        let create_command = format!(
            r#"$process = Start-Process 'sc.exe' -ArgumentList 'create {} binPath= "cmd.exe /c \"{}\"" start= auto' -Verb RunAs -PassThru; $process.WaitForExit(); Write-Output $process.ExitCode"#,
            format!("{}", self.service_name),
            bat_file_path
        );

        match run_powershell_command_with_output(&create_command) {
            Ok(output) => {
                let re = Regex::new(r"\d+").unwrap();
                let numbers: Vec<i32> = re
                    .find_iter(&output)
                    .filter_map(|digits| digits.as_str().parse::<i32>().ok())
                    .collect();
                let output_code: i32 = numbers[0];

                if output_code == 0 {
                    // success
                    execute!(
                        stdout,
                        cursor::MoveTo(0, message_row as u16),
                        SetForegroundColor(Color::Green),
                        Print("✓ Служба успешно установлена.\n"),
                        ResetColor
                    )?;
                } else if output_code == 5 {
                    // access denied
                    execute!(
                        stdout,
                        cursor::MoveTo(0, message_row as u16),
                        SetForegroundColor(Color::Yellow),
                        Print("⚠ В доступе отказано, не удалось установить службу.\n"),
                        ResetColor
                    )?;
                } else if output_code == 740 {
                    // elevation required
                    execute!(
                        stdout,
                        cursor::MoveTo(0, message_row as u16),
                        SetForegroundColor(Color::Yellow),
                        Print("⚠ Требуется повышение прав, не удалось установить службу.\n"),
                        ResetColor
                    )?;
                } else if output_code == 1073 {
                    // service already installed
                    execute!(
                        stdout,
                        cursor::MoveTo(0, message_row as u16),
                        SetForegroundColor(Color::Yellow),
                        Print("⚠ Служба уже установлена.\n"),
                        ResetColor
                    )?;
                } else {
                    execute!(
                        stdout,
                        cursor::MoveTo(0, message_row as u16),
                        SetForegroundColor(Color::Red),
                        Print(format!("⚠ Ошибка при установке службы: {}\n", output_code)),
                        ResetColor
                    )?;
                }
                message_row += 2;
            },
            Err(e) => {
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Red),
                    Print(format!("⚠ Ошибка при установке службы: {}\n", e)),
                    ResetColor
                )?;
                message_row += 2;
            }
        }

        Ok(message_row)
    }

    fn start_service(&self, stdout: &mut impl Write, mut message_row: usize) -> io::Result<usize> {
        execute!(
            stdout,
            cursor::MoveTo(0, message_row as u16),
            Print("► Запуск службы...\n")
        )?;
        message_row += 1;
        stdout.flush()?;

        let command = format!(
            "Start-Process 'sc.exe' -ArgumentList 'start {}' -Verb RunAs",
            self.service_name
        );

        match run_powershell_command_with_output(&command) {
            Ok(_) => {
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Green),
                    Print("✓ Служба успешно запущена.\n"),
                    ResetColor
                )?;
                message_row += 2; // Add extra line for spacing
            },
            Err(e) => {
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Red),
                    Print(format!("⚠ Ошибка при запуске службы: {}\n", e)),
                    ResetColor
                )?;
                message_row += 2; // Add extra line for spacing
            }
        }

        Ok(message_row)
    }
}