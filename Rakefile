require 'fileutils'
require 'rake/clean'
require 'zip'

CLEAN.include('*.tmp', '*.log', '*.rar', '*.zip')
CLOBBER.include('*.exe')

DIST_FILENAME = "dist.zip"
EXE_FILENAME = "mikaelly.exe"

namespace 'exe' do
  desc "Create exe file for windows machines"
  task :create do
    FileUtils.rm_rf EXE_FILENAME
    `ocra main.rbw ProcessadorDePendencias.rb --output #{EXE_FILENAME} --add-all-core --gemfile Gemfile --no-dep-run --gem-full`
  end
end

desc "Create .zip file with files for distribution"
task :dist do
  input_filenames = %w(main.rbw ProcessadorDePendencias.rb Rakefile README.md LICENSE Gemfile mikaelly.exe)
  FileUtils.rm_rf DIST_FILENAME
  Zip::File.open(DIST_FILENAME, Zip::File::CREATE) do |zipfile|
    input_filenames.each do |filename|
      zipfile.add(filename, filename)
    end
  end
end