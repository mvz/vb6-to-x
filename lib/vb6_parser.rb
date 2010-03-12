require 'rubygems'
require 'treetop'
require 'vb6'

module VB6ToX
  class VB6Parser
    def parse_or_abort data
      result = self.parse data
      if result.nil?
	raise self.failure_reason
      end
      result
    end
  end
end
