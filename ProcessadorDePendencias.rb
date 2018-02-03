class String
  def is_integer?
    self =~ /^[-+]?([0-9]*)?$/
  end

  def remover_acentuacao
    self.tr( "ÀÁÂÃÄÅàáâãäåĀāĂăĄąÇçĆćĈĉĊċČčÐðĎďĐđÈÉÊËèéêëĒēĔĕĖėĘęĚěĜĝĞğĠġĢģĤĥĦħÌÍÎÏìíîïĨĩĪīĬĭĮįİıĴĵĶķĸĹĺĻļĽľĿŀŁłÑñŃńŅņŇňŉŊŋÒÓÔÕÖØòóôõöøŌōŎŏŐőŔŕŖŗŘřŚśŜŝŞşŠšſŢţŤťŦŧÙÚÛÜùúûüŨũŪūŬŭŮůŰűŲųŴŵÝýÿŶŷŸŹźŻżŽž",
             "AAAAAAaaaaaaAaAaAaCcCcCcCcCcDdDdDdEEEEeeeeEeEeEeEeEeGgGgGgGgHhHhIIIIiiiiIiIiIiIiIiJjKkkLlLlLlLlLlNnNnNnNnnNnOOOOOOooooooOoOoOoRrRrRrSsSsSsSssTtTtTtUUUUuuuuUuUuUuUuUuUuWwYyyYyYZzZzZz")
  end
end

module ProcessadorDePendencias
  require 'tiny_tds'
  require 'concurrent'
  require 'connection_pool'
  
  BANK_TO_SPREADSHEET_COLUMN_NUMBER_OF_PROPOSALS = {
   #"bank_name"  => ["proposal_column", "typer_column"],
    "teste" => "E",
    "help"  => "E",
    "ole"   => "H",
    "itau"  => "I",
    "intermed emprest" => "A",
    "intermed cart" => "A",
    "daycoval" => "T",
    "cetelem" => "A",
    #"ccb" => "",
    "bradesco" => "A",
    "bons" => "D",
    "safra" => "B",
    "sabemi" => "C",
    "pan consignado" => "A",
    "banrisul" => "F",
    "pan cartao" => "A"
  }
  
  def openSpreadSheet filename
    if filename.end_with?("xls")
      require "roo-xls"
    else
      require "roo"
    end
    
    Roo::Spreadsheet.open(filename)
  end
  
  def getSpreadSheetColumnNames filename
    spreadsheet = openSpreadSheet filename
    column_names = spreadsheet.row(0)
    spreadsheet.close
    return column_names
  end
  
  def getSpreadSheetColumn filename, bank
    bank = bank.remover_acentuacao.downcase
    bank_keys = BANK_TO_SPREADSHEET_COLUMN_NUMBER_OF_PROPOSALS.keys
    if bank_keys.include? bank
      puts "Bank: #{bank}"
      return BANK_TO_SPREADSHEET_COLUMN_NUMBER_OF_PROPOSALS[bank]
    else
      file_basename = File.basename(filename).remover_acentuacao.downcase
      bank_keys.each do |k|
        if file_basename.include? k 
          puts "Bank: #{k}"
          return BANK_TO_SPREADSHEET_COLUMN_NUMBER_OF_PROPOSALS[k]
        end
      end
    end
    raise "O banco desta planilha não pôde ser encontrado"
  end
  
  def acquireListOfProposalsAndColumns file, bank
    proposals_column_number = getSpreadSheetColumn file, bank
    
    spreadsheet = openSpreadSheet file
    proposals = spreadsheet.column(proposals_column_number)
    other_columns = (spreadsheet.first_column..spreadsheet.last_column).each_with_object(Array.new) do |column_idx, column_list|
      column_list << spreadsheet.column(column_idx) unless column_idx == (proposals_column_number.ord - "A".ord + 1) #converts proposals_column_number to integer index of column"
    end
    spreadsheet.close
    columns = proposals.zip(other_columns.transpose)
    columns.each_with_object(Array.new) do |(p, t), columns_array|
      if p.is_a? Integer
        columns_array << [p.to_s, t]
      elsif p.is_a? Float
        columns_array << [p.to_s.slice(0..-3), t] #Removes ".0"
      elsif p.is_a? String
        if p =~ /[0-9]{2}\-[0-9]{9}\/[0-9]{2}/ #CETELEM
          columns_array << [p[3..11], t]
        elsif p =~ /[0-9]*\-[0-9]/ #Bradesco
          columns_array << [p[1..-3], t]
        elsif p.is_integer?
          columns_array << [p.to_i.to_s, t] #Remove trailing zeroes (BANRISUL)
        else #Header line
          tag_regexp = /(<[^>]*>)|\n|\t/
          columns_array << [p.gsub(tag_regexp, " ").split.join(" "), t.collect{|h| h.to_s.gsub(tag_regexp, " ").split.join(" ")}]
        end
      end
    end
  end
  
  def createDatabaseConnection
    begin
      TinyTds::Client.new username: user_login, password: user_passwd, host: database_url, database: database_name, timeout: 10*60
    rescue TinyTds::Error => err
      raise "Falha ao se conectar ao banco: " + err.message
    end 
  end
  
  def getSQL proposal
    "SELECT UE.SGL_UNIDADE_EMPRESA, UE.SGL_UNIDADE_FEDERACAO, UE.NOM_UNIDADE_EMPRESA, UE.NOM_FANTASIA
      FROM [CBDATA].[dbo].[PROPOSTA_EMPRESTIMO] AS PE
        INNER JOIN [CBDATA].[dbo].[UNIDADE_EMPRESA] AS UE ON UE.COD_UNIDADE_EMPRESA = PE.COD_UNIDADE_EMPRESA
      WHERE PE.NUM_PROPOSTA = '#{proposal}' OR PE.NUM_CONTRATO = '#{proposal}'"
  end
  
  def queryDatabaseForProposal con, proposal
    sql = getSQL proposal
    result = con.execute sql
    row = result.first
    result.cancel
    if row.nil?
      nil
    else
      uf = row["SGL_UNIDADE_EMPRESA"]
      uf = uf.nil? ? row["SGL_UNIDADE_FEDERACAO"] : uf
      nome_loja = row["NOM_UNIDADE_EMPRESA"]
      nome_loja = nome_loja.nil? ? row["NOM_FANTASIA"] : nome_loja
      return [uf, nome_loja]
    end
  end
  
  def findUfOfEachProposal(full_data, progress_keeper=nil, num_connections=15)
    thread_pool = Concurrent::FixedThreadPool.new(2*num_connections, fallback_policy: :discard)
    con_pool = ConnectionPool.new(size: num_connections, timeout: 10*60) { createDatabaseConnection }
    begin
      raise "Não foi possível acessar o banco de dados" unless con_pool
      mutex = Mutex.new
      
      threadLog = Logger.new STDOUT
      threadLog.formatter = proc do |severity, datetime, progname, msg|
        "Thread #{Thread.current.object_id} | #{msg}\n"
      end
      
      failed_proposals = Concurrent::Array.new
      response = Concurrent::Array.new
      
      header_not_captured = true
      full_data.each do |proposal_number, other_columns|
        if proposal_number.is_integer?
          thread_pool.post do
            uf = nil
            nome_loja = nil
            con_pool.with do |con|
              uf, nome_loja = queryDatabaseForProposal con, proposal_number
            end
            
            if uf
              other_columns.insert(0, nome_loja)
              other_columns.insert(0, uf)
              other_columns.insert(0, proposal_number)
              response << other_columns
            else
              failed_proposals << proposal_number unless failed_proposals.include?(proposal_number)
            end
            mutex.synchronize {progress_keeper.progress += 1 unless progress_keeper.nil?}
          end
        else
          other_columns.insert(0, "Nome loja Workbank")
          other_columns.insert(0, "Workbank UF")
          other_columns.insert(0, "Proposta/Contrato")
          response << other_columns
          mutex.synchronize do 
            progress_keeper.progress += 1 unless progress_keeper.nil?
            raise "Erro ao encontrar propostas... \n Altere o nome da planilha ou selecione o banco correto acima" unless header_not_captured
            header_not_captured = false
          end
        end
      end
    ensure
      puts "Waiting processing of proposals"
      
      thread_pool.shutdown
      thread_pool.wait_for_termination
      puts "Threads finished"
      con_pool.shutdown { |con| con.close }
      puts "Connections closed"
    end
    return [response.collect {|i| i.flatten}, failed_proposals]
  end
  
  def recoverProposalNumbersAndStateOfProposals(file, bank, progress_keeper=nil)
    full_data = acquireListOfProposalsAndColumns file, bank
    if full_data.one? == 1
      raise "Nenhuma proposta localizada"
    else
      puts "Total number of proposals: #{full_data.length - 1}" #Remove header line
    end
    progress_keeper.total = full_data.length unless progress_keeper.nil?
    findUfOfEachProposal full_data, progress_keeper
  end
end

if __FILE__ == $0 ### script de teste ###
  include ProcessadorDePendencias
  
  abort "Passe apenas o caminho do arquivo excel como parametro" unless ARGV.length == 1
  recoverProposalNumbersAndStateOfProposal(ARGV[0], "teste")
end