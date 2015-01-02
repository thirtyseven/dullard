require 'dullard'

describe "test.xlsx," do
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
          row[0].strftime("%Y-%m-%d %H:%M:%S").should == "2012-10-18 00:00:00"
          row[1].strftime("%Y-%m-%d %H:%M:%S").should == "2012-10-18 00:17:58"
          row[2].strftime("%Y-%m-%d %H:%M:%S").should == "2012-07-01 21:18:48"
          row[3].strftime("%Y-%m-%d %H:%M:%S").should == "2012-07-01 21:18:52"
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
          row[0].should == 'teststring'
          row[1].strftime("%Y-%m-%d %H:%M:%S").should == "2012-10-18 00:00:00"
          row[2].strftime("%Y-%m-%d %H:%M:%S").should == "2012-10-18 00:17:58"
          row[3].strftime("%Y-%m-%d %H:%M:%S").should == "2012-07-01 21:18:48"
          row[4].strftime("%Y-%m-%d %H:%M:%S").should == "2012-07-01 21:18:52"
        end
      end
      count.should == 117
    end
  end
end

describe "test2.xlsx" do
  before(:each) do
    @file = File.open(File.expand_path("../test2.xlsx", __FILE__))
  end

  it "should not skip nils" do
    rows = Dullard::Workbook.new(@file).sheets[0].rows.to_a
    rows.should == [
      [1],
      [nil, 2],
      [nil, nil, 3]
    ]
  end
end
