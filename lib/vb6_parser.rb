require 'rubygems'
require 'treetop'
require 'vb6'
require 'tire_swing'

module VB6ToX
  module AST
    include TireSwing::NodeDefinition
    node :root, :version, :layout
    node :version, :value
  end

  TireSwing.parses_grammar(VB6, AST)

  include TireSwing::VisitorDefinition

  visitor :array_visitor do
    visits AST::Root do |root|
      arr = []
      visit(root.version, arr) unless root.version == ""
      arr
    end
    visits AST::Version do |version, arr|
      arr << [:version, version.value]
    end
  end
end
