require File.dirname(__FILE__) + '/test_helper'

class WholeFileTest < Test::Unit::TestCase
  include VB6ToX
  context "Parsing a whole file" do
    setup do
      @parser = VB6Parser.new
      @data = File.read(File.dirname(__FILE__) + "/files/randtext.frm")
    end
    should "work" do
      ast = nil
      assert_nothing_raised do
	ast = @parser.parse_or_abort(@data)
      end
      assert_not_nil ast
    end
  end
end

