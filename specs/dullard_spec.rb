require 'dullard'

describe "dullard" do
  before(:each) do
    @file = File.open("SHAPE5_CorePrePost.xlsx")
    @xlsx = Dullard::Workbook.new @file
  end
  it "can open a file" do
    @xlsx.should_not be_nil
  end

  it "can find sheets" do
    @xlsx.sheets.count.should == 1
  end

  it "can read rows" do
    @xlsx.sheets[0].rows.first.count.should >= 300
  end

  it "reads the right number of rows" do
    count = 0
    @xlsx.sheets[0].rows.each do |row|
      count += 1
    end
    count.should == 115
  end

  it "reads the right number of rows from the metadata" do
    @xlsx.sheets[0].rows.size.should == 115
  end
end
