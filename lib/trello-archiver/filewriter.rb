
module TrelloArchiver

  class Authorize
      include Trello
      include Trello::Authorization

      def initialize(config)
        @config = config
      end

      def authorize
        Trello::Authorization.const_set :AuthPolicy, OAuthPolicy

        credential = OAuthCredential.new(@config['public_key'], @config['private_key'])
        OAuthPolicy.consumer_credential = credential

        OAuthPolicy.token = OAuthCredential.new( @config['access_token_key'], nil )
      end
  end

  class Prompt
    def initialize(config)
      @config = config
    end

    def get_board

      me = Member.find("me")
      boardarray = Array.new
      optionnum = 1
      me.boards.each do |board|
        boardarray << board
        puts "#{optionnum}: #{board.name} #{board.id}"
        optionnum += 1
      end

      puts "0 - CANCEL\n\n"
      puts "Which board would you like to backup?"
      if @config['board'].nil?
        board_to_archive = gets.to_i - 1
      else
        board_to_archive = @config['board'] - 1
      end

      if board_to_archive == -1
         puts "Cancelling"
         exit 1 
      end

      board = Board.find(boardarray[board_to_archive].id)
    end

    def get_filename

      puts "Would you like to provide a filename? (y/n)"

      if @config['filename'] == 'default'
        filename = @board.name.parameterize
      else
        response = gets.downcase.chomp
         if response.to_s =~ /^y/i
           puts "Enter filename:"
           filename = gets
         else
           filename = @board.name.parameterize
         end
      end


        puts "Preparing to backup #{@board.name}"
        lists = @board.lists
        filename
    end

    def run
      @board = get_board
      @filename = get_filename
      result = {}
      result[:board] = @board
      result[:filename] = @filename
      result
    end
  end

  class Archiver
    def initialize(options = {:board => "", :filename => "trello_backup", :format => 'xlsx', :col_sep => ","})
      @options = options
      FileUtils.mkdir("archive") unless Dir.exists?("archive")
      date = DateTime.now.strftime "%Y%m%dT%H%M"
      @filename = "#{Dir.pwd}/#{date}_#{@options[:filename].upcase}.#{@options[:format]}"
    end

    def create_backup()
      case @options[:format]
      when 'csv' && ( @options[:col_sep] == "\t" )
        @options[:format] = 'tsv'
        create_csv
      when 'tsv'
        @options[:col_sep] = "\t"
        create_csv
      when 'csv'
        create_csv
      when 'xlsx'
        create_xlsx
      else
        #
        puts "Trello-archiver can create csv and xlsx backups. Please choose one of these options and try again."
      end
      
    end

    def create_csv()
      require 'CSV'
      # Filename= filename or default of boardname
      # 
      # Board object has been passed into the method
      lists = @options[:board].lists

      CSV.open(@filename, "w", :col_sep => @options[:col_sep]) do |csv|
        # Sheets have to be restructured as a field
        #
        # Add header row
        csv << %w[Name Description Labels Progress Comments]

        lists.each do |list|
          puts list.name
          cards = list.cards
          
          cards.each do |card|
            # Add title row
            puts card.name
            # gather and join the labels if they exist
            labels = case card.labels.length
            when 0
              "none"
            else
              card.labels.map { |c| c.name }.join(" ")
            end

            # Gather comments
            comments = card.actions.map do |action|
              if action.type == "commentCard"
                # require 'pry'; binding.pry
                "#{Member.find(action.member_creator_id).full_name} [#{ action.date.strftime('%m/%d/%Y') }] : #{action.data['text']} \n\n"
              end
            end

            csv << [card.name, card.description, labels, list.name, comments.join('')]
          end
        end
      end
    end

    def create_xlsx()
      require 'xlsx_writer'
      # Filename= filename or default of boardname
      # 
      # Board object has been passed into the method
      lists = @options[:board].lists

      @doc = XlsxWriter.new

      lists.each do |list|
        sheet = @doc.add_sheet(list.name.gsub(/\W+/, '_'))
        puts list.name
        cards = list.cards
        #
        # Add header row
        sheet.add_row( %w[Name Description Labels Comments])
        
        cards.each do |card|
          # Add title row
          puts card.name
          # sheet.add_row([
          #   "Title: #{card.name}",
          #   "Desc: #{card.description}",
          #   "Labels: #{card.labels.length}"
          # ])
          
          # gather and join the labels if they exist
          labels = case card.labels.length
          when 0
            "none"
          else
            card.labels.map { |c| c.name }.join(" ")
          end

          # Gather comments
          comments = card.actions.map do |action|
            if action.type == "commentCard"
              # require 'pry'; binding.pry
              "#{Member.find(action.member_creator_id).full_name} [#{ action.date.strftime('%m/%d/%Y') }] : #{action.data['text']} \n\n"
            end
          end

          sheet.add_row([card.name, card.description, labels, comments.join('')])
        end
      end

      # Moving file to where I want it
      require 'fileutils'
      ::FileUtils.mv @doc.path, @filename

      # Cleanup of temp dir
      @doc.cleanup
    end

  end
end
