require 'time' unless defined? Time 
require 'date' unless defined? Date 
require 'datetime' unless defined? DateTime

class Float

  def from_excel
    seconds_in_day ||= 24.0 * 60.0 * 60.0
    excel_days_offset ||= 25569.0
    gmt_offset = (((Time.now.hour.to_f - Time.now.gmtime.hour.to_f)/24.0) * -1) + (0.5 / seconds_in_day) # Have to compensate for GMT offsets since Excel does by default, and then we add half a second to make sure it has the correct day (and not 23:59:59 of the previous day)
    val = self
    val += gmt_offset
    new_date = Time.at((val - excel_days_offset) * seconds_in_day)
  end
  alias_method :from_xl, :from_excel

end


class Integer

  def from_excel
    self.to_f.from_excel
  end
  alias_method :from_xl, :from_excel

end


class Fixnum

    def from_excel
      self.to_f.from_excel
    end
    alias_method :from_xl, :from_excel
end


