RubyExcel
=========

An attempt to create Excel-like Workbooks in Ruby

Example
-------

```ruby

def create_sample_data
  a=[];8.times{|t|b=[];c='A';5.times{b<<"#{c}#{t+1}";c.next!};a<<b};a
end

require 'RubyExcel'
re = RubyExcel.new

s = re.load create_sample_data
puts s

s.rows(3,4).each { |r| puts r['A'] }

puts s.range 'B2:C4'

s.range( s.cell(1,1), s.cell(2,2) ).delete
puts s

s.column(3).delete
s.column('D').each { |el| puts el }

s2 = re.add 'NewSheet'
s2.load create_sample_data.transpose
rng = s2.range 'A1:B3'
rng.value = rng.map { |cell| cell + 'a' }
	
```