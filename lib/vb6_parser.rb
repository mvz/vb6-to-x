require 'rubygems'
require 'treetop'
require 'vb6'
require 'tire_swing'

module AST
  include TireSwing::NodeDefinition
  node :vb6, :version, :layout
  node :versionspec, :version
end

#Treetop.load(File.join(File.dirname(__FILE__), "vb6.treetop"))
TireSwing.parses_grammar(VB6, AST)


