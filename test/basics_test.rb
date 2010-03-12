require File.dirname(__FILE__) + '/test_helper'

class BasicsTest < Test::Unit::TestCase
  include VB6ToX
  context "The VB6 parser" do
    setup do
      @parser = VB6Parser.new
    end
    should "parse version" do
      ast = @parser.parse("VERSION 5.00")
      assert_equal "5.00", ast.version.value.text_value
      assert ast.version.value.terminal?
    end
    should "parse empty string" do
      ast = @parser.parse("")
      assert_equal "", ast.text_value
    end
    context "using the parse_or_abort method" do
      should "raise an error when the parse fails" do
	assert_nil @parser.parse "foo"
	assert_raises RuntimeError do
	  @parser.parse_or_abort "foo"
	end
      end
    end
  end
end
