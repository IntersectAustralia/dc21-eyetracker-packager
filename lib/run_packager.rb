require File.expand_path(File.dirname(__FILE__) + '/logger')
require 'win32ole'
require 'yaml'
require 'mimetype_fu'
require 'archive/zip'

arg_dest = ARGV[0]
arg_source = ARGV[1]

interactive = arg_source.nil?

log_file_path = File.join(File.dirname(__FILE__), '..', 'log', "eyeTRackerToDIVER-#{Time.now.strftime("%Y-%m-%d")}-log.txt")
log_writer = Logger.new(log_file_path)

log_writer.log_message('INFO', 'EyeTracker packager starting...')

if arg_dest
  transfer_path = arg_dest
else
  puts "Please enter the DIVER transfer pending path (eg. C:\\Pending):"
  transfer_path = STDIN.gets
  transfer_path.strip!
end

if !Dir.exists?(transfer_path.to_s)
  log_writer.log_message('ERROR','You need to specify a valid DIVER transfer pending folder')
  log_writer.log_message('INFO', 'Aborting...')
  log_writer.close
  exit
end

def calculate_filename(source_folder, filename)
  original = File.join(source_folder, filename).gsub(/\\/, "/")
  return  original unless File.exists?(original)
  ext = File.extname(original).to_s

  regex = if ext.eql?("")
            name = original
            /\A#{Regexp.escape(name)}_(\d+)\Z/
          else
            name = original[0..(original.rindex(".") - 1)]
            /\A#{Regexp.escape(name)}_(\d+)\.#{Regexp.escape(ext[1..-1])}\Z/
          end
  matching = Dir[File.join(name + "*")].collect do |s|
    match = s.match(regex)
    match ? match[1].to_i : nil
  end
  numbers = matching.compact.sort
  next_number = ((1..numbers[-1].to_i+1).to_a - numbers)[0]

  name + "_#{next_number}" + ext
end


def extract_sessions(transfer_path, log_writer, interactive = false)
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  available_drives = []
  fso.Drives.each do |drive|
    if drive.IsReady and !drive.VolumeName.to_s.eql?("")
      display = "#{drive.DriveLetter}:"
      display << " (#{drive.VolumeName})" 
      available_drives <<  display
    end
  end

  puts "Detected the following drives with volume labels connected:"
  available_drives.each do |d|
    puts "- #{d}"
  end
  puts ""

  if interactive
    puts "Please enter the drive letter or path you wish to import"
    puts "or press Ctrl + C twice to exit:"
    source_path = STDIN.gets.strip
  else
    source_path = ARGV[1].dup
  end

  #convert backslashes to slash
  source_path.gsub!(/\\/, "/")

  if source_path[/^\w$/]
    source_path = source_path + ':/'
  elsif source_path[/^\w:$/]
    source_path = source_path + '/'
  end

  if !Dir.exists?(source_path)
    log_writer.log_message('ERROR',"The path #{source_path} does not exist.")
    log_writer.log_message('INFO', 'Aborting...')
    puts ""
    if interactive
      return
    else
      log_writer.close
      exit
    end
  end

  # if it's just the drive, get the volume label
  if source_path[/^\w:\/$/]
    label = fso.getDrive(source_path).VolumeName
    if label.to_s.strip.eql?("")
      log_writer.log_message('ERROR','The drive specified does not have a volume label.')
      log_writer.log_message('INFO', 'Aborting...')
      puts ""
      if interactive
        return
      else
        log_writer.close
        exit
      end
    end
  else
    label = File.basename(source_path)
  end

  log_writer.log_message('INFO',"Importing #{source_path} to #{transfer_path}...")
  log_writer.log_message('INFO',"Using #{label} as volume label...")
  puts ""

  root_logs = Dir[File.join(source_path,"*.log")].sort
  root_sessions = Dir[File.join(source_path,"Session_*")].sort_by {|e| e.split(/(\d+)/).map {|a| a =~ /\d+/ ? a.to_i : a }}
  root_remainder = Dir[File.join(source_path, '*')] - root_logs - root_sessions
  root_images = root_remainder.select{|a| File.mime_type?(a)[/^image/]}
  root_others = root_remainder - root_images

  if root_logs.empty? and root_sessions.empty?
    log_writer.log_message('ERROR', "No Eyetracker Session folders or logs detected in #{source_path}.")
    log_writer.log_message('INFO', 'Aborting...')
    puts ""

    if interactive
      return
    else
      log_writer.close
      exit
    end
  end

  ###### Root Logs
  log_writer.log_message('INFO',"Zipping EyeTracker logs...")
  if root_logs.empty?
    log_writer.log_message('INFO',"  No EyeTracker logs found; Ignoring.")
  else
    log_dates = root_logs.collect{|log| File.ctime(log).strftime("%Y-%m-%d")}.uniq

    log_dates.each do |log_date|
      log_writer.log_message('INFO',"  Zipping EyeTracker logs for #{log_date}...")
      root_log_zip = calculate_filename(transfer_path,"eT-SD#{label}-LOGS-#{log_date.gsub('-', '')}.zip")
      Archive::Zip.archive(root_log_zip, Dir[File.join(source_path, log_date + "*.log")])
      log_writer.log_message('INFO',"    Created #{root_log_zip}.")
    end
    puts ""
  end
  ### End Root Logs

  ###### Root Images
  log_writer.log_message('INFO',"Zipping EyeTracker images...")
  if root_images.empty?
    log_writer.log_message('INFO',"  No EyeTracker images found; Ignoring.")
  else
    root_image_hash = {}
    root_images.each do |image|
      created_at = File.ctime(image).strftime("%Y%m%d")
      root_image_hash[created_at] ||= []
      root_image_hash[created_at] << image
    end
    root_image_hash.each do |cdate,files|
      log_writer.log_message('INFO',"  Zipping EyeTracker images for #{cdate}...")
      root_images_zip = calculate_filename(transfer_path,"eT-SD#{label}-IMAGES-#{cdate}.zip")
      Archive::Zip.archive(root_images_zip, files)
      log_writer.log_message('INFO',"    Created #{root_images_zip}.")
    end
  end

  puts ""
  ### End Root Images

  ###### Root Others
  log_writer.log_message('INFO',"Zipping EyeTracker others...")
  if root_others.empty?
    log_writer.log_message('INFO',"  No EyeTracker others found; Ignoring.")
  else
    root_others_cdate = File.ctime(root_logs[0] || root_others[0]).strftime("%Y%m%d")
    root_others_zip = calculate_filename(transfer_path,"eT-SD#{label}-OTHERS-#{root_others_cdate}.zip")
    Archive::Zip.archive(root_others_zip, root_others)
    log_writer.log_message('INFO',"    Created #{root_others_zip}.")

  end
  puts ""

  ### End Root Others

  ###### EyeTracker sessions
  log_writer.log_message('INFO',"Processing EyeTracker sessions...")
  puts ""

  root_sessions.each do |session_folder|
    log_writer.log_message('INFO',"  Processing EyeTracker #{File.basename(session_folder)}...")
    session_all = Dir[File.join(session_folder,"*.*")].sort

    session_glasses = Dir[File.join(session_folder,"*.{log,dbg,raw,dat}")].sort
    session_remainder = session_all - session_glasses

    session_images = session_remainder.select{|a| File.mime_type?(a)[/^image/]}
    session_audio = session_remainder.select{|a| File.mime_type?(a)[/^audio/]}
    session_others = session_remainder - session_audio - session_images

    session_cdate = File.ctime(session_folder).strftime("%Y%m%d")

    # glasses
    if session_glasses.empty?
      log_writer.log_message('INFO',"    #{File.basename(session_folder)} does not contain EyeTracker files; Ignoring.")
    else
      transfer_name = "eT-SD#{label}-#{File.basename(session_folder)}-GLASSES-#{session_cdate}.zip"
      session_glasses_zip = calculate_filename(transfer_path, transfer_name)
      Archive::Zip.archive(session_glasses_zip, session_glasses)
      log_writer.log_message('INFO',"    Created #{session_glasses_zip}.")
    end

    #audio
    if session_audio.empty?
      log_writer.log_message('INFO',"    #{File.basename(session_folder)} does not contain audio files; Ignoring.")
    else
      log_writer.log_message('INFO',"    Copying EyeTracker #{File.basename(session_folder)} audio files...")
      session_audio.each do |audio|
        filename = File.basename(audio)
        extension = File.extname(filename)
        audio_cdate = File.ctime(audio).strftime("%Y%m%d")

        transfer_name =  filename.gsub(/^/, "eT-SD#{label}-#{File.basename(session_folder)}-INTERVIEW-").sub(extension, "-#{audio_cdate}#{extension}")
        new_path = calculate_filename(transfer_path,transfer_name)
        FileUtils.cp audio, new_path
        log_writer.log_message('INFO',"      Copied #{filename}.")
      end
    end

    #image
    if session_audio.empty?
      log_writer.log_message('INFO',"    #{File.basename(session_folder)} does not contain image files; Ignoring.")
    else
      log_writer.log_message('INFO',"    Copying EyeTracker #{File.basename(session_folder)} image files...")
      session_images.each do |image|
        filename = File.basename(image)
        extension = File.extname(filename)
        image_cdate = File.ctime(image).strftime("%Y%m%d")

        transfer_name =  filename.gsub(/^/, "eT-SD#{label}-#{File.basename(session_folder)}-DOCKET-").sub(extension, "-#{image_cdate}#{extension}")
        new_path = calculate_filename(transfer_path,transfer_name)
        FileUtils.cp image, new_path
        log_writer.log_message('INFO',"      Copied #{filename}.")
      end
    end

    #others
    if session_others.empty?
      log_writer.log_message('INFO',"    #{File.basename(session_folder)} does not contain other files; Ignoring.")
    else
      session_others_zip = calculate_filename(transfer_path,"eT-SD#{label}-#{File.basename(session_folder)}-OTHERS-#{session_cdate}.zip")
      Archive::Zip.archive(session_others_zip, session_others)
      log_writer.log_message('INFO',"    Created #{session_others_zip}.")
    end
    log_writer.log_message('INFO',"  #{File.basename(session_folder)} processed.")
  end
  
  log_writer.log_message('INFO',"Imported #{source_path} to #{transfer_path}.")
  puts ""

  log_writer.flush_current(calculate_filename(transfer_path,"eT-SD#{label}-Manifest.txt"))

  if interactive
    puts "Please insert a new card if required and then press Enter to continue."
    puts "If not, press Ctrl + C twice to exit."
    unused = STDIN.gets
  end

end

if interactive
  while true
    extract_sessions(transfer_path, log_writer, true)
  end
else
  extract_sessions(transfer_path, log_writer)

end

