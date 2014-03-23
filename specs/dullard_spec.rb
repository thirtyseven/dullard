require 'dullard'

describe "dullard," do
  before(:each) do
    @file = File.open(File.expand_path("../test.xlsx", __FILE__))
  end

  describe "when it has no user defined formats," do
    before(:each) do
      @xlsx = Dullard::Workbook.new @file
    end

    it "can open a file" do
      @xlsx.should_not be_nil
    end

    it "can find sheets" do
      @xlsx.sheets.count.should == 1
    end

    it "reads the right number of columns, even with blanks" do
      rows = @xlsx.sheets[0].rows
      rows.next.count.should == 300
      rows.next.count.should == 9
      rows.next.count.should == 1
    end

    it "reads the right number of rows" do
      @xlsx.sheets[0].row_count.should == 117
    end

    it "reads the right number of rows from the metadata when present" do
      @xlsx.sheets[0].rows.size.should == 117
    end

    it "reads date/time properly" do
      count = 0
      @xlsx.sheets[0].rows.each do |row|
        count += 1

        if count == 116
          row.should == ["2012-10-20 00:00:00", "2012-10-20 00:17:58", "2012-07-03 21:18:48", "2012-07-03 21:18:52"]
        end
      end
      count.should == 117
    end
  end

  describe "when it has user defined formats," do
    before(:each) do
      @xlsx = Dullard::Workbook.new @file, {'GENERAL' => :string, 'm/d/yyyy' => :date, 'M/D/YYYY' => :date,}
    end

    it "converts the user defined formatted cells properly" do
      count = 0
      @xlsx.sheets[0].rows.each do |row|
        count += 1

        if count == 117
          row.should == ["teststring", "2012-10-20 00:00:00", "2012-10-20 00:17:58", "2012-07-03 21:18:48", "2012-07-03 21:18:52"]
        end
      end
      count.should == 117
    end
  end
end
