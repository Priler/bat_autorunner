use std::io::{self, Write};
use crossterm::{
    cursor,
    execute,
    style::{Print, SetForegroundColor, Color, ResetColor},
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

        match run_powershell_command(&command) {
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

        match run_powershell_command(command) {
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

        match run_powershell_command(&command) {
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

        let command = format!(
            "Start-Process 'sc.exe' -ArgumentList 'create {} binPath= \"cmd.exe /c \"\"{}\"\"\" start= auto' -Verb RunAs",
            self.service_name,
            bat_file_path
        );

        match run_powershell_command(&command) {
            Ok(_) => {
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Green),
                    Print("✓ Служба успешно установлена.\n"),
                    ResetColor
                )?;
                message_row += 2; // Add extra line for spacing
            },
            Err(e) => {
                execute!(
                    stdout,
                    cursor::MoveTo(0, message_row as u16),
                    SetForegroundColor(Color::Red),
                    Print(format!("⚠ Ошибка при установке службы: {}\n", e)),
                    ResetColor
                )?;
                message_row += 2; // Add extra line for spacing
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

        match run_powershell_command(&command) {
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