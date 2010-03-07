require 'rake'
require 'rake/testtask'

file 'lib/vb6.rb' => ['lib/vb6.treetop'] do |t|
  sh "tt #{t.prerequisites.join} -o #{t.name}"
end

Rake::TestTask.new do |test|
  test.pattern = 'test/*_test.rb'
  test.verbose = true
end

task :test => ['lib/vb6.rb'] 
