require 'rubygems'
require 'treetop'
require 'vb6'
require 'tire_swing'

module VB6ToX
  module AST
    include TireSwing::NodeDefinition
    node :root, :version, :layout
    node :version, :version
  end

  TireSwing.parses_grammar(VB6, AST)
end
