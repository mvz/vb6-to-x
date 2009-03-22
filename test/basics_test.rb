require File.dirname(__FILE__) + '/test_helper'

class BasicsTest < Test::Unit::TestCase
  context "The VB6 parser" do
    setup do
      @parser = VB6Parser.new
    end
    should "parse version" do
      assert_equal [], VB6Parser.ast("VERSION 5.00")
    end
    should "parse empty string" do
      assert_equal [], VB6Parser.ast("")
    end
  end
end
