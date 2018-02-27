class String
  def is_integer?
     /^[0-9]+$/ =~ self.strip
  end

  def blank?
    self.empty?
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
    "primeira coluna" => "A",
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
    
    begin
      Roo::Spreadsheet.open(filename)
    rescue StandardError
      raise "Erro ao abrir a planilha, feche a planilha caso ela já esteja aberta"
    end
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
    else # Deteção automática
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
          columns_array << [p.gsub(tag_regexp, " ").split.join(" "), t.collect{|h| h.to_s.gsub(tag_regexp, " ").split.join(" ")}] #Remover tags html
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
  
  def getSqlFor proposal
    "SELECT top(2) UE.SGL_UNIDADE_EMPRESA, UE.SGL_UNIDADE_FEDERACAO, UE.NOM_UNIDADE_EMPRESA, UE.NOM_FANTASIA
      FROM [CBDATA].[dbo].[VW_PROPOSTA_CONSULTA] AS PE
        INNER JOIN [CBDATA].[dbo].[UNIDADE_EMPRESA] AS UE ON UE.COD_UNIDADE_EMPRESA = PE.COD_UNIDADE_EMPRESA
      WHERE PE.PROPOSTA = '#{proposal}' OR PE.CONTRATO = '#{proposal}'"
  end
  
  def queryDatabaseForProposal con, proposal
    sql = getSqlFor proposal
    result = con.execute sql
    rows = result.each
    return nil unless rows.one?
    result.cancel
    
    row = rows.first
    if row.nil?
      nil
    else
      nome_loja = row["NOM_UNIDADE_EMPRESA"]
      if nome_loja.nil? then nome_loja = row["NOM_FANTASIA"] end
      
      uf = row["SGL_UNIDADE_EMPRESA"]
      if uf.nil? then uf = row["SGL_UNIDADE_FEDERACAO"] end
      
      ufs_brasil = %w(AC AL AM AP BA CE DF ES GO MA MG MS MT PA PB PE PI PR RJ RN RO RR RS SC SE SP TO)
      uf_from_nome_loja = nome_loja.to_s.split.first
      if uf.nil? and ufs_brasil.include?(uf_from_nome_loja) then uf = uf_from_nome_loja end
      return [uf, nome_loja]
    end
  end
  
  def findUfOfEachProposal(full_data, progress_keeper=nil, num_connections=30)
    con_pool = ConnectionPool.new(size: num_connections, timeout: 10*60) { createDatabaseConnection }
    thread_pool = Concurrent::FixedThreadPool.new(2*num_connections, fallback_policy: :discard)
    failed_proposals = Concurrent::Array.new
    response = Concurrent::Array.new
    mutex = Mutex.new
    
    begin 
      header_not_captured = true
      full_data.each do |proposal_number, other_columns|
        if proposal_number.is_integer?
          thread_pool.post do
            uf = nome_loja = nil
            con_pool.with do |con|
              uf, nome_loja = queryDatabaseForProposal con, proposal_number
            end
            if uf or nome_loja
              other_columns.insert(0, proposal_number)
              other_columns.insert(1, uf)
              other_columns.insert(2, nome_loja)
              response.push other_columns
            else
              failed_proposals.push proposal_number unless failed_proposals.include? proposal_number
            end
            mutex.synchronize {progress_keeper.progress += 1} unless progress_keeper.nil?
          end
        else
          next if proposal_number.strip == "-" #Blank proposal
          other_columns.insert(0, "Proposta/Contrato")
          other_columns.insert(1, "UF Workbank")
          other_columns.insert(2, "Loja Workbank")
          response.push other_columns
          mutex.synchronize do 
            progress_keeper.progress += 1 unless progress_keeper.nil?
            raise "Erro ao encontrar propostas... Altere o nome da planilha ou selecione o banco correto acima" unless header_not_captured
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