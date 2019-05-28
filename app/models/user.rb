class User

  include ActiveModel::Conversion
  extend ActiveModel::Naming
  attr_accessor :id, :name, :email

  def initialize
  end

  def initialize id=nil, name=nil, email=nil
    @id = id
    @name = name
    @email = email
  end

  def save
    #open file
    workbook = RubyXL::Parser.parse(Rails.root.join("db", "users.xlsx").to_s)
    worksheet = workbook[0]
    @users = []
    flag = true
    index= 1
    while flag
      if worksheet[index] #workbook[100] returns nil if no value
        
        index = index + 1
      else
        if index == 1
          @id = 1
        else
          @id = worksheet[index-1][0].value + 1
        end
        row = worksheet.insert_row(index)
        worksheet[index][0].change_contents(@id, worksheet[index][0].formula)
        worksheet[index][1].change_contents(@name, worksheet[index][1].formula)
        worksheet[index][2].change_contents(@email, worksheet[index][2].formula)
        flag = false

        #save file
        workbook.write(Rails.root.join("db", "users.xlsx").to_s)
        return true
      end
    end
    false
  end

  def update id, name, email
    workbook = RubyXL::Parser.parse(Rails.root.join("db", "users.xlsx").to_s)
    worksheet = workbook[0]
    flag = true
    index= 1
    while flag
      if worksheet[index] #workbook[100] returns nil if no value
        if worksheet[index][0].value == @id.to_i
          worksheet[index][1].change_contents(name, worksheet[index][1].formula)
          worksheet[index][2].change_contents(email, worksheet[index][2].formula)
          workbook.write(Rails.root.join("db", "users.xlsx").to_s)
          return true
        end
        
        index = index + 1
      else
        flag = false
      end
    end
    false
  end


  def destroy
    workbook = RubyXL::Parser.parse(Rails.root.join("db", "users.xlsx").to_s)
    worksheet = workbook[0]
    flag = true
    index= 1
    while flag
      if worksheet[index] #workbook[100] returns nil if no value
        if worksheet[index][0].value == @id.to_i
          row = worksheet.delete_row(index)
          workbook.write(Rails.root.join("db", "users.xlsx").to_s)
          puts "Id is ----------------------------------------- #{@id}"
          return true
        end
        
        index = index + 1
      else
        flag = false
      end
    end
    puts "----------------------------------------------------Shit"
    false
  end

  class << self
    def all
      workbook = RubyXL::Parser.parse(Rails.root.join("db", "users.xlsx").to_s)
      worksheet = workbook[0]
      @users = []
      flag = true
      index= 1
      while flag
        if worksheet[index] #workbook[100][100] returns nil if no record
          user = User.new worksheet[index][0].value, worksheet[index][1].value, worksheet[index][2].value
          @users << user
          index = index + 1
        else
          flag = false
        end
      end
      @users
    end

    def find id
      workbook = RubyXL::Parser.parse(Rails.root.join("db", "users.xlsx").to_s)
      worksheet = workbook[0]
      flag = true
      index= 1
      @user
      while flag
        if worksheet[index] #workbook[100][100] returns nil if no record
          if worksheet[index][0].value == id.to_i
            @user = User.new worksheet[index][0].value, worksheet[index][1].value, worksheet[index][2].value
            flag = false
          end
        else
          flag = false
        end
        index = index + 1
      end
      @user
    end
  end
end
