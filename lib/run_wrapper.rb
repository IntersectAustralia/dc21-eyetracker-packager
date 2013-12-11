require File.expand_path(File.dirname(__FILE__) + '/api_call_logger')
require 'win32ole'
require 'yaml'
require 'mimetype_fu'
require 'archive/zip'

arg_source = ARGV[0]
arg_dest = ARGV[1]

interactive = arg_source.nil? && arg_dest.nil?

log_file_path = File.join(File.dirname(__FILE__), '..', 'log', "eyeTRackerToDIVER-#{Time.now.strftime("%Y-%m-%d")}-log.txt")
log_writer = ApiCallLogger.new(log_file_path)

log_writer.log_message('INFO', 'EyeTracker packager starting...')

puts "Please enter the DIVER transfer pending path (eg. C:\\Users\\Guest\\Desktop\\Pending):"
# transfer_path = STDIN.gets
#TODOD
transfer_path = "C:/TRANSFER"
transfer_path.strip!

if !Dir.exists?(transfer_path.to_s)
  log_writer.log_message('ERROR','You need to specify a valid DIVER transfer pending folder')
  log_writer.log_message('INFO', 'Aborting...')
  log_writer.close
  exit
  # print error message here
  # next
end

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

# while true
  puts "Please enter the drive letter or path you wish to import"
  puts "or press Ctrl + C twice to exit:"
  source_path = STDIN.gets
  #convert backslashes to slash
  source_path.strip!.gsub!(/\\/, "/")

  if source_path[/^\w$/]
    source_path = source_path + ':'
  end

  if !Dir.exists?(source_path)
    log_writer.log_message('ERROR','You need to pass an argument containing the drive or path you want to import')
    log_writer.log_message('INFO', 'Aborting...')
    # log_writer.close
    # exit
    # print error message here
    # next
  end

  # if it's just the drive, get the volume label
  if source_path[/^\w:$/]
    #do drive label stuff
    label = fso.getDrive(source_path).VolumeName
  else
    #use folder name
    label = File.basename(source_path)
  end

  log_writer.log_message('INFO',"Importing #{source_path} to #{transfer_path}...")
  log_writer.log_message('INFO',"Using #{label} as volume label...")

  root_logs = Dir[File.join(source_path,"*.log")].sort
  root_sessions = Dir[File.join(source_path,"Session_*")].sort
  root_remainder = Dir[File.join(source_path, '*')] - root_logs - root_sessions
  root_images = root_remainder.select{|a| File.mime_type?(a)[/^image/]}
  root_others = root_remainder - root_images

  ###### Root Logs
  log_writer.log_message('INFO',"Zipping EyeTracker logs...")

  log_dates = root_logs.collect{|log| File.ctime(log).strftime("%Y-%m-%d")}.uniq

  log_dates.each do |log_date|
    log_writer.log_message('INFO',"  Zipping EyeTracker logs for #{log_date}...")
    root_log_zip = File.join(transfer_path,"eT-SD#{label}-LOGS-#{log_date.gsub('-', '')}.zip")
    Archive::Zip.archive(root_log_zip, Dir[File.join(source_path, log_date + "*.log")])
  end

  ### End Root Logs

  ###### Root Images
  log_writer.log_message('INFO',"Zipping EyeTracker images...")
  root_image_hash = {}
  root_images.each do |image|
    created_at = File.ctime(image).strftime("%Y%m%d")
    root_image_hash[created_at] ||= []
    root_image_hash[created_at] << image
  end

  root_image_hash.each do |cdate,files|
      log_writer.log_message('INFO',"  Zipping EyeTracker images for #{cdate}...")
      root_images_zip = File.join(transfer_path,"eT-SD#{label}-IMAGES-#{cdate}.zip")
      Archive::Zip.archive(root_images_zip, files)
  end
  ### End Root Images

  ###### Root Others
  log_writer.log_message('INFO',"Zipping EyeTracker others...")
  root_others_cdate = File.ctime(root_logs[0]).strftime("%Y%m%d")
  root_others_zip = File.join(transfer_path,"eT-SD#{label}-OTHERS-#{root_others_cdate}.zip")
  Archive::Zip.archive(root_others_zip, root_others)
  ### End Root Others

  ###### EyeTracker sessions
  log_writer.log_message('INFO',"Zipping EyeTracker sessions etc...")

  root_sessions.each do |session_folder|
  log_writer.log_message('INFO',"  Zipping EyeTracker sessions etc for #{File.basename(session_folder)}...")
    session_all = Dir[File.join(session_folder,"*.*")].sort

    session_glasses = Dir[File.join(session_folder,"*.{log,dbg,raw,dat}")].sort
    session_remainder = session_all - session_glasses

    session_images = session_remainder.select{|a| File.mime_type?(a)[/^image/]}
    session_audio = session_remainder.select{|a| File.mime_type?(a)[/^audio/]}
    session_others = session_remainder - session_audio - session_images

    session_cdate = File.ctime(session_folder).strftime("%Y%m%d")
    # glasses
    session_glasses_zip = File.join(transfer_path,"eT-SD#{label}-#{File.basename(session_folder)}-GLASSES-#{session_cdate}.zip")
    Archive::Zip.archive(session_glasses_zip, session_glasses)

    #audio
    session_audio.each do |audio|
      filename = File.basename(audio)
      extension = File.extname(filename)
      audio_cdate = File.ctime(audio).strftime("%Y%m%d")

      transfer_name =  filename.gsub(/^/, "eT-SD#{label}-#{File.basename(session_folder)}-INTERVIEW-").sub(extension, "-#{audio_cdate}#{extension}")
      FileUtils.cp audio, File.join(transfer_path,transfer_name)
    end

    #image
    session_images.each do |image|
      filename = File.basename(image)
      extension = File.extname(filename)
      image_cdate = File.ctime(image).strftime("%Y%m%d")

      transfer_name =  filename.gsub(/^/, "eT-SD#{label}-#{File.basename(session_folder)}-DOCKET-").sub(extension, "-#{image_cdate}#{extension}")
      FileUtils.cp image, File.join(transfer_path,transfer_name)
    end

    #others
    session_others_zip = File.join(transfer_path,"eT-SD#{label}-#{File.basename(session_folder)}-OTHERS-#{session_cdate}.zip")

    Archive::Zip.archive(session_others_zip, session_others)

  end
# end