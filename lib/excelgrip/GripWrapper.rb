module Excelgrip
  class GripWrapper
    def initialize(raw_object)
      if raw_object.methods.include?("raw")
        @raw_object = raw_object.raw
      else
        @raw_object = raw_object
      end
    end
    
    def raw
      @raw_object
    end
    
    def inspect()
      "<#{self.class}>"
    end
    
    private
    
    OLE_METHODS = [:Type, :Initialize]
    def method_missing(m_id, *params)
      unless OLE_METHODS.include?(m_id)
        missing_method_name = m_id.to_s.downcase
        methods.each {|method|
          if method.to_s.downcase == missing_method_name
            return send(method, *params)
          end
        }
      end
      # Undefined Method is throwed to raw_object
      begin
        @raw_object.send(m_id, *params)
      rescue
        raise $!,$!.message, caller
      end
    end
    
    def toRaw(target_obj)
      target_obj.respond_to?("raw") ? target_obj.raw : target_obj
    end
  end
  
  class FileSystemObject
    include Singleton
    def initialize
      @body =  WIN32OLE.new('Scripting.FileSystemObject')
    end
    
    def fullpath(filename)
      @body.getAbsolutePathName(filename)
    end
  end
end
