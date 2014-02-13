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
  log_writer.log_message('ERROR', 'You need to specify a valid DIVER transfer pending folder')
  log_writer.log_message('INFO', 'Aborting...')
  log_writer.close
  exit
end

def calculate_filename(source_folder, filename)
  original = File.join(source_folder, filename).gsub(/\\/, "/")
  return original unless File.exists?(original)

  @counts[:duplicate] += 1
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

  @counts = {duplicate: 0,
            errors: {},
            processed: 0,
            created: 0,
            audio: 0,
            images: 0
  }

  fso = WIN32OLE.new('Scripting.FileSystemObject')
  available_drives = []
  fso.Drives.each do |drive|
    if drive.IsReady and !drive.VolumeName.to_s.eql?("")
      display = "#{drive.DriveLetter}:"
      display << " (#{drive.VolumeName})"
      available_drives << display
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
    original_source = STDIN.gets.strip
  else
    original_source = ARGV[1].dup
  end

  #convert backslashes to slash
  original_source.gsub!(/\\/, "/")

  if original_source[/^\w$/]
    original_source = original_source + ':/'
  elsif original_source[/^\w:$/]
    original_source = original_source + '/'
  end

  # check if path exists
  if !Dir.exists?(original_source)
    log_writer.log_message('ERROR', "The path #{original_source} does not exist.")
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
  if original_source[/^\w:\/$/]
    label = fso.getDrive(original_source).VolumeName
    if label.to_s.strip.eql?("")
      log_writer.log_message('ERROR', 'The drive specified does not have a volume label.')
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
    label = File.basename(original_source)
  end

  log_writer.log_message('INFO', "Importing #{original_source} to #{transfer_path}...")

  # Copy files onto computer first
  log_writer.log_message('INFO', "  Copying to temporary folder, this may take up to a few minutes...")
  puts ""

  package_tmp_path = File.expand_path(File.join(File.dirname(__FILE__), '..', 'tmp', label))


  # keeps copying files until the directory mirrors
  cmd = "robocopy \"#{original_source}\" \"#{package_tmp_path}\" /XD .* /XF .* /ZB /MIR /COPY:DAT /XA:SHT /E /DCOPY:T /R:10 /W:5 /NP /LOG+:robocopy.log /TEE"

  puts ""
  output = []
  IO.popen(cmd).each do |line|
    puts line
    output << line
  end

  robocopy_summary = output[-10..-1].select { |a| a[/(?:Files|Dir)/] }

  robocopy_summary.each do |line|
    components = line.strip.split(/\s+/)
    @counts[:errors][components[0].downcase.to_sym] = components[6].to_i
  end

  log_writer.log_message('INFO', "Using #{label} as volume label...")
  puts ""


  root_projects = Dir[File.join(package_tmp_path, '*')].select { |e| File.directory?(e) && File.basename(e).size <= 20 && !Dir[File.join(e, "Session_*")].empty? }
  root_others = Dir[File.join(package_tmp_path, '*')] - root_projects

  ###### Root Others
  log_writer.log_message('INFO', "Zipping root others...")
  if root_others.empty?
    log_writer.log_message('INFO', "  No root others found; Ignoring.")
  else
    root_others_zip = calculate_filename(transfer_path, "eT-SD#{label}-OTHERS.zip")
    Archive::Zip.archive(root_others_zip, root_others)
    @counts[:created] += 1
    log_writer.log_message('INFO', "    Created #{root_others_zip} with:")
    root_others.each do |file|
      log_writer.log_message('INFO', "      - #{file.gsub(package_tmp_path + '/', "")}")
    end
  end
  puts ""

  ### End Root Others

  ### Process Each Project
  log_writer.log_message('INFO', "Processing projects with session folders...")
  puts ""

  root_projects.each do |source_path|

    project = File.basename(source_path)

    log_writer.log_message('INFO', "Processing #{project}...")
    puts ""

    project_logs = Dir[File.join(source_path, "*.log")].select do |log|
      !File.basename(log)[/^\d{4}-\d{2}-\d{2}__\d\d_\d\d_\d\d\.log$/].nil?
    end

    project_sessions = Dir[File.join(source_path, "Session_*")].sort_by { |e| e.split(/(\d+)/).map { |a| a =~ /\d+/ ? a.to_i : a } }
    project_remainder = Dir[File.join(source_path, '*')] - project_logs - project_sessions
    project_images = project_remainder.select { |a| File.mime_type?(a)[/^image/] }
    project_others = project_remainder - project_images

    ###### Project Logs
    log_writer.log_message('INFO', "  Zipping #{project} logs...")
    if project_logs.empty?
      log_writer.log_message('INFO', "    No #{project} logs found; Ignoring.")
    else

      log_writer.log_message('INFO', "    Zipping #{project} logs...")
      project_log_zip = calculate_filename(transfer_path, "eT-SD#{label}-#{project}-LOGS.zip")

      files = project_logs
      Archive::Zip.archive(project_log_zip, files)
      @counts[:created] += 1

      log_writer.log_message('INFO', "      Created #{project_log_zip} with:")
      files.each do |file|
        @counts[:processed] += 1
        log_writer.log_message('INFO', "        - #{file.gsub(source_path + '/', "")}")
      end

    end

    puts ""
    ### End Project Logs

    ###### Project Images
    log_writer.log_message('INFO', "  Zipping #{project} images...")
    if project_images.empty?
      log_writer.log_message('INFO', "    No #{project} images found; Ignoring.")
    else
      project_image_hash = {}
      project_images.each do |image|
        created_at = File.ctime(image).strftime("%Y%m%d")
        project_image_hash[created_at] ||= []
        project_image_hash[created_at] << image
      end
      project_image_hash.each do |cdate, files|
        log_writer.log_message('INFO', "    Zipping #{project} images for #{cdate}...")
        project_images_zip = calculate_filename(transfer_path, "eT-SD#{label}-#{project}-IMAGES-#{cdate}.zip")
        Archive::Zip.archive(project_images_zip, files)
        @counts[:created] += 1
        log_writer.log_message('INFO', "      Created #{project_images_zip} with:")
        files.each do |file|
          @counts[:processed] += 1
          log_writer.log_message('INFO', "        - #{file.gsub(source_path + '/', "")}")
        end
      end
    end

    puts ""
    ### End Project Images

    ###### Project Others
    log_writer.log_message('INFO', "  Zipping #{project} others...")
    if project_others.empty?
      log_writer.log_message('INFO', "    No #{project} others found; Ignoring.")
    else
      project_others_zip = calculate_filename(transfer_path, "eT-SD#{label}-#{project}-OTHERS.zip")
      Archive::Zip.archive(project_others_zip, project_others)
      @counts[:created] += 1
      log_writer.log_message('INFO', "      Created #{project_others_zip} with:")
      project_others.each do |file|
        @counts[:processed] += 1
        log_writer.log_message('INFO', "        - #{file.gsub(source_path + '/', "")}")
      end
    end
    puts ""

    ### End Project Others

    ###### EyeTracker sessions
    log_writer.log_message('INFO', "  Processing #{project} sessions...")

    if project_sessions.empty?
      log_writer.log_message('INFO', "    No #{project} sessions found; Ignoring.")
    else
      project_sessions.each do |session_folder|
        log_writer.log_message('INFO', "    Processing #{project} #{File.basename(session_folder)}...")
        session_all = Dir[File.join(session_folder, "*.*")].sort

        session_glasses = Dir[File.join(session_folder, "*.{log,dbg,raw,dat}")].sort
        session_remainder = session_all - session_glasses

        session_images = session_remainder.select { |a| !File.mime_type?(a)[/^image/].nil? }
        session_audio = session_remainder.select { |a| !File.mime_type?(a)[/^audio/].nil? }
        session_others = session_remainder - session_audio - session_images

        session_cdate = File.ctime(session_folder).strftime("%Y%m%d")

        # glasses
        if session_glasses.empty?
          log_writer.log_message('INFO', "      #{File.basename(session_folder)} does not contain EyeTracker files; Ignoring.")
        else
          transfer_name = "eT-SD#{label}-#{project}-#{File.basename(session_folder)}-GLASSES-#{session_cdate}.zip"
          session_glasses_zip = calculate_filename(transfer_path, transfer_name)
          Archive::Zip.archive(session_glasses_zip, session_glasses)
          @counts[:created] += 1
          log_writer.log_message('INFO', "      Created #{session_glasses_zip} with:")
          session_glasses.each do |file|
            @counts[:processed] += 1
            log_writer.log_message('INFO', "        - #{file.gsub(source_path + '/', "")}")
          end
        end

        #audio
        if session_audio.empty?
          log_writer.log_message('INFO', "      #{File.basename(session_folder)} does not contain audio files; Ignoring.")
        else
          log_writer.log_message('INFO', "      Copying #{project} #{File.basename(session_folder)} audio files...")
          session_audio.each do |audio|
            filename = File.basename(audio)
            extension = File.extname(filename)
            audio_cdate = File.ctime(audio).strftime("%Y%m%d")

            transfer_name = filename.gsub(/^/, "eT-SD#{label}-#{project}-#{File.basename(session_folder)}-INTERVIEW-").sub(extension, "-#{audio_cdate}#{extension}")
            new_path = calculate_filename(transfer_path, transfer_name)
            FileUtils.cp audio, new_path
            @counts[:created] += 1
            @counts[:processed] += 1
            @counts[:audio] += 1
            log_writer.log_message('INFO', "        - #{audio.gsub(source_path + '/', "")}")
          end
        end

        #image
        if session_images.empty?
          log_writer.log_message('INFO', "      #{File.basename(session_folder)} does not contain image files; Ignoring.")
        else
          log_writer.log_message('INFO', "      Copying #{project} #{File.basename(session_folder)} image files...")
          session_images.each do |image|
            filename = File.basename(image)
            extension = File.extname(filename)
            image_cdate = File.ctime(image).strftime("%Y%m%d")

            transfer_name = filename.gsub(/^/, "eT-SD#{label}-#{project}-#{File.basename(session_folder)}-DOCKET-").sub(extension, "-#{image_cdate}#{extension}")
            new_path = calculate_filename(transfer_path, transfer_name)
            FileUtils.cp image, new_path
            @counts[:created] += 1
            @counts[:processed] += 1
            @counts[:images] += 1
            log_writer.log_message('INFO', "        - #{image.gsub(source_path + '/', "")}")
          end
        end

        #others
        if session_others.empty?
          log_writer.log_message('INFO', "      #{File.basename(session_folder)} does not contain other files; Ignoring.")
        else
          session_others_zip = calculate_filename(transfer_path, "eT-SD#{label}-#{project}-#{File.basename(session_folder)}-OTHERS-#{session_cdate}.zip")
          Archive::Zip.archive(session_others_zip, session_others)
          @counts[:created] += 1
          log_writer.log_message('INFO', "      Created #{session_others_zip} with:")
          session_others.each do |file|
            @counts[:processed] += 1
            log_writer.log_message('INFO', "        - #{file.gsub(source_path + '/', "")}")
          end
        end
        log_writer.log_message('INFO', "    #{project} #{File.basename(session_folder)} processed.")
        puts ""

      end
    end
    puts ""

  end

  if @counts[:errors][:dirs] == 0 && @counts[:errors][:files] == 0
    # Remove robocopied files
    FileUtils.rm_rf package_tmp_path
    log_writer.log_message('INFO', "Cleaned temporary folder.")
  else
    log_writer.log_message('WARN', "Errors found during copy, files will be kept in #{package_tmp_path}")
  end
  puts ""

  manifest_path = calculate_filename(transfer_path, "eT-SD#{label}-Manifest-#{Time.now.strftime("%Y%m%d")}.txt")
  @counts[:created] += 1

  log_writer.log_message('INFO', "Summary (including manifest file):")
  log_writer.log_message('INFO', "  Number of files failed to copy:   #{@counts[:errors][:files]}")
  log_writer.log_message('INFO', "  Number of folders failed to copy: #{@counts[:errors][:dirs]}")
  log_writer.log_message('INFO', "  Number of files processed:        #{@counts[:processed]}")
  log_writer.log_message('INFO', "  Number of files created:          #{@counts[:created]}")
  log_writer.log_message('INFO', "    - Number of docket files:       #{@counts[:images]}")
  log_writer.log_message('INFO', "    - Number of interview files:    #{@counts[:audio]}")
  log_writer.log_message('INFO', "    - Number of duplicated files:   #{@counts[:duplicate]}")
  puts ""

  log_writer.log_message('INFO', "Imported #{original_source} to #{transfer_path}.")
  puts ""

  log_writer.flush_current(manifest_path)

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

