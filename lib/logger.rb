require 'json'

class Logger
  ENTRY_DELIMITER = '-------------------------------------------------------'
  attr_accessor :log_file, :current_output

  def initialize(file_path)
    self.log_file = File.open(file_path, 'a')
    self.current_output = Tempfile.new(Time.now.strftime("%Y%m%d%H%M%S"))
  end

  def flush_current(path)
    current_output.close
    FileUtils.cp current_output.path, path
    self.current_output = Tempfile.new(Time.now.strftime("%Y%m%d%H%M%S"))
  end

  def log_message(severity, message)
    puts message

    time = timestamp

    log_file.printf time
    log_file.printf 'ERROR ' if severity == 'ERROR'
    log_file.printf 'INFO  ' if severity == 'INFO'
    log_file.printf 'WARN  ' if severity == 'WARN'
    log_file.printf '' if severity == 'SYM'
    log_file.puts message

    current_output.printf time
    current_output.printf 'ERROR ' if severity == 'ERROR'
    current_output.printf 'INFO  ' if severity == 'INFO'
    current_output.printf 'WARN  ' if severity == 'WARN'
    current_output.printf '' if severity == 'SYM'
    current_output.puts message
  end

  def close
    log_file.puts ENTRY_DELIMITER
    log_file.close
    current_output.close

  end

  def timestamp
    "#{Time.now} "
  end


end
