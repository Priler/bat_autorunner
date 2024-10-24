use std::io::{self, Write};
use crossterm::{
    cursor,
    execute,
    style::{Print, SetForegroundColor, Color, ResetColor},
    terminal::{Clear, ClearType},
};
use crate::utils::run_powershell_command;

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
        // Initial message
        execute!(
            stdout,
            cursor::MoveTo(0, message_row as u16),
            Clear(ClearType::FromCursorDown),
            Print("=== Удаление существующей службы ===")
        )?;
        message_row += 1;
        stdout.flush()?;

        // Stop service
        message_row = self.stop_service(stdout, message_row)?;

        // Terminate process
        message_row = self.terminate_process(stdout, message_row)?;

        // Delete service
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
            Print(format!("► Установка файла как службы: {}", bat_file_path))
        )?;
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
            Print(format!("► Остановка службы '{}'...", self.service_name))
        )?;

        let command = format!(
            "Start-Process 'sc.exe' -ArgumentList 'stop {}' -Verb RunAs",
            self.service_name
        );

        match run_powershell_command(&command) {
            Ok(_) => {
                message_row += 1;
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Green),
                    Print(format!("✓ Служба '{}' успешно остановлена.", self.service_name)),
                    ResetColor
                )?;
            },
            Err(e) => {
                message_row += 1;
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Red),
                    Print(format!("⚠ Ошибка при остановке службы: {}", e)),
                    ResetColor
                )?;
            }
        }

        Ok(message_row + 1)
    }

    fn terminate_process(&self, stdout: &mut impl Write, mut message_row: usize) -> io::Result<usize> {
        execute!(
            stdout,
            cursor::MoveTo(0, message_row as u16),
            Print("► Завершение процесса 'winws.exe'...")
        )?;
        stdout.flush()?;

        let command = "Start-Process 'powershell' -ArgumentList 'Stop-Process -Name \"winws\" -Force' -Verb RunAs";

        match run_powershell_command(command) {
            Ok(_) => {
                message_row += 1;
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Green),
                    Print("✓ Процесс 'winws.exe' успешно завершён."),
                    ResetColor
                )?;
            },
            Err(e) => {
                message_row += 1;
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Red),
                    Print(format!("⚠ Ошибка при завершении процесса 'winws.exe': {}", e)),
                    ResetColor
                )?;
            }
        }

        Ok(message_row + 1)
    }

    fn delete_service(&self, stdout: &mut impl Write, mut message_row: usize) -> io::Result<usize> {
        execute!(
            stdout,
            cursor::MoveTo(0, message_row as u16),
            Print(format!("► Удаление службы '{}'...", self.service_name))
        )?;
        stdout.flush()?;

        let command = format!(
            "Start-Process 'sc.exe' -ArgumentList 'delete {}' -Verb RunAs",
            self.service_name
        );

        match run_powershell_command(&command) {
            Ok(_) => {
                message_row += 1;
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Green),
                    Print(format!("✓ Служба '{}' успешно удалена.", self.service_name)),
                    ResetColor
                )?;
            },
            Err(e) => {
                message_row += 1;
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Red),
                    Print(format!("⚠ Ошибка при удалении службы: {}", e)),
                    ResetColor
                )?;
            }
        }

        Ok(message_row + 1)
    }

    fn create_service(&self, stdout: &mut impl Write, bat_file_path: &str, mut message_row: usize) -> io::Result<usize> {
        execute!(
            stdout,
            cursor::MoveTo(0, message_row as u16),
            Print("► Создание службы...")
        )?;
        stdout.flush()?;

        let command = format!(
            "Start-Process 'sc.exe' -ArgumentList 'create {} binPath= \"cmd.exe /c \"\"{}\"\"\" start= auto' -Verb RunAs",
            self.service_name,
            bat_file_path
        );

        match run_powershell_command(&command) {
            Ok(_) => {
                message_row += 1;
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Green),
                    Print("✓ Служба успешно установлена."),
                    ResetColor
                )?;
            },
            Err(e) => {
                message_row += 1;
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Red),
                    Print(format!("⚠ Ошибка при установке службы: {}", e)),
                    ResetColor
                )?;
            }
        }

        Ok(message_row + 1)
    }

    fn start_service(&self, stdout: &mut impl Write, mut message_row: usize) -> io::Result<usize> {
        execute!(
            stdout,
            cursor::MoveTo(0, message_row as u16),
            Print("► Запуск службы...")
        )?;
        stdout.flush()?;

        let command = format!(
            "Start-Process 'sc.exe' -ArgumentList 'start {}' -Verb RunAs",
            self.service_name
        );

        match run_powershell_command(&command) {
            Ok(_) => {
                message_row += 1;
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Green),
                    Print("✓ Служба успешно запущена."),
                    ResetColor
                )?;
            },
            Err(e) => {
                message_row += 1;
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Red),
                    Print(format!("⚠ Ошибка при запуске службы: {}", e)),
                    ResetColor
                )?;
            }
        }

        Ok(message_row + 1)
    }
}