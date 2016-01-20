# to_xl

Small gem to convert ruby Time, Date, and DateTime instances to Excel date integers. This is helpful when using some Excel parsers to compare the dates in the spreadsheet to dates you may be using elsewhere in the script. 

# Usage
Simple! Require the gem `require 'to_xl', initialize a Time, Date or DateTime value. Call `.to_xl` on it.

# Example
```ruby
irb(main):002:0> require 'to_xl'
=> true
irb(main):003:0> Time.now.to_xl
=> 42389
irb(main):004:0> Date.today.to_xl
=> 42389
irb(main):005:0> DateTime.now.to_xl
=> 42389
```
##### That result in Excel:
![The above integer in Excel (and formated)](http://i.imgur.com/LcUIWro.png "In Excel")
