require 'time'
gem_dir = Gem::Specification.find_by_name("to_xl").gem_dir
load "#{gem_dir}/lib/from_xl.rb"

class Time

  def to_excel
    @seconds_in_day ||= 24 * 60 * 60
    @excel_days_offset ||= 25569
    @offset_gmt = (self.hour > self.gmtime.hour) ? true : false  # Excel dates are on GMT time, so we have to go back a day if this is calculated on one day in our time zone and another on GMT.
    excel_date = (self.to_i/@seconds_in_day) + @excel_days_offset
    excel_date -= 1 if @offset_gmt # Excel dates are on GMT time, so we have to go back a day if this is calculated on one day in our time zone and another on GMT.
    excel_date
  end
  alias_method :to_xl, :to_excel

  def say_hi
    puts "Hello there"
  end

end


if defined?(Date)
  class Date

    def to_excel
      self.to_time.to_excel
    end
    alias_method :to_xl, :to_excel

  end
end


if defined?(DateTime)
  class DateTime

    def to_excel
      self.to_time.to_excel
    end
    alias_method :to_xl, :to_excel
    
  end
end