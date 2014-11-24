begin
  require 'win32ole' unless Module.const_defined?(:WIN32OLE)
rescue LoadError
end
require 'test/unit'

if Module.const_defined?(:WIN32OLE)
  class TestThread < Test::Unit::TestCase
    #
    # test for Bug #2618(ruby-core:27634)
    #
    def assert_creating_win32ole_object_in_thread(meth)
      skip "mruby-thread is required to test this case" unless Module.const_defined?(:Thread)
      skip "Thread##{meth} method not found" unless Thread.respond_to?(meth)
      t = Thread.__send__(meth) {
        ex = nil
        begin 
          WIN32OLE.new('Scripting.Dictionary')
        rescue => ex
        end
        assert_equal(nil, ex, "[Bug #2618] Thread.#{meth}")
      }
      t.join
    end

    def test_creating_win32ole_object_in_thread_new
      assert_creating_win32ole_object_in_thread(:new)
    end

    def test_creating_win32ole_object_in_thread_start
      assert_creating_win32ole_object_in_thread(:start)
    end

    def test_creating_win32ole_object_in_thread_fork
      assert_creating_win32ole_object_in_thread(:fork)
    end
  end
end
