require File.dirname(__FILE__) + '/test_helper'

class BasicsTest < Test::Unit::TestCase
  include VB6ToX
  context "The VB6 parser" do
    setup do
      @parser = VB6Parser.new
    end
    should "parse version" do
      assert_equal [[:version, "5.00"]], parse_tree_array("VERSION 5.00")
    end
    should "parse empty string" do
      assert_equal [], parse_tree_array("")
    end
  end

  def parse_tree_array(s)
    ArrayVisitor.visit(VB6Parser.ast(s))
  end
end
