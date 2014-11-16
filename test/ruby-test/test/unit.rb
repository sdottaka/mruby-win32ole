
module Assertions
  def assert_instance_of(cls, obj)
  end
  def assert_raise(e)
  end
  def assert_equal(e, r)
  end
  def assert_nil(o)
  end
  def assert_in_delta(e, r, d)
  end
  def assert_match(er, r)
  end
end

module Test
  module Unit
    class TestCase
      include Assertions
    end
  end
end

# vim: set ts=2 sw=2 expandtab:
