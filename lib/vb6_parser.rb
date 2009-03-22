require 'rubygems'
require 'treetop'
require 'vb6'
require 'tire_swing'

module AST
  include TireSwing::NodeDefinition
  node :versionspec, :version
end

#Treetop.load("tire_swing.treetop")
TireSwing.parses_grammar(VB6, AST)


