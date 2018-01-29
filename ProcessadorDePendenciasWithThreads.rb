class String
  def is_integer?
    self =~ /^[-+]?[1-9]([0-9]*)?$/
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
    "teste" => ["E", "F"],
    "help"  => ["E", "J"],
    "ole"   => ["H", "D"],
    "itau"  => ["I", "E"],
    "intermed emprest" => ["A", "N"],
    "intermed card" => ["A", "M"],
    "daycoval" => ["T", "S"],
    "centelem" => ["", ""],
    "ccb" => ["", ""],
    "bradesco" => ["A", "J"],
    "bons" => ["D", "A"],
    "safra" => ["B", "S"],
    "sabemi" => ["C", "A"],
    "pan consignado" => ["A", "Q"],
    "pan cartao" => ["A", "P"],
    "banrisul" => ["", ""]
  }
  
  def getSpreadSheetColumn filename, bank
    bank = bank.remover_acentuacao.downcase
    bank_keys = BANK_TO_SPREADSHEET_COLUMN_NUMBER_OF_PROPOSALS.keys
    if bank_keys.include? bank
      return BANK_TO_SPREADSHEET_COLUMN_NUMBER_OF_PROPOSALS[bank]
    else
      file_basename = File.basename(filename).remover_acentuacao.downcase
      bank_keys.each do |k|
        if file_basename.include? k 
          return BANK_TO_SPREADSHEET_COLUMN_NUMBER_OF_PROPOSALS[k]
        end
      end
    end
    raise "O banco desta planilha não pôde ser encontrado"
  end
  
  def acquireListOfProposals file, bank
    if file.end_with?("xls")
      require "roo-xls"
    else
      require "roo"
    end
    
    spreadsheet =  Roo::Spreadsheet.open(file)
    proposals_column_number, typer_column_number = getSpreadSheetColumn file, bank
    
    proposals = spreadsheet.column(proposals_column_number)
    typers = spreadsheet.column(typer_column_number)
    columns = proposals.zip(typers)
    columns.reject {|p, t| p.nil?}.select {|p, t| p.is_a?(Integer) or p.is_a?(Float) or p.is_integer?}.collect {|p, t| p.is_a?(Float) ? [Integer(p), t] : [p, t]}
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
  
  def queryDatabaseUfOfProposal con, proposal
    sql = getSQL proposal
    result = con.execute sql
    puts result.inspect unless result.is_a? TinyTds::Result
    row = result.first
    result.cancel
    if row.nil?
      nil
    else
      uf = row["SGL_UNIDADE_EMPRESA"]
      uf = uf.nil? ? row["SGL_UNIDADE_FEDERACAO"] : uf
      uf = uf.nil? ? row["NOM_UNIDADE_EMPRESA"] : uf
      uf = uf.nil? ? row["NOM_FANTASIA"] : uf
    end
  end
  
  def findUfOfEachProposal(proposals_and_typers, progress_keeper=nil, num_connections=15)
    begin
      con_pool = ConnectionPool.new(size: num_connections, timeout: 10*60) { createDatabaseConnection }
      raise "Não foi possível acessar o banco de dados" unless con_pool
      mutex = Mutex.new
      
      threadLog = Logger.new STDOUT
      threadLog.formatter = proc do |severity, datetime, progname, msg|
        "Thread #{Thread.current.object_id} | #{msg}\n"
      end
      
      failed_proposals = Concurrent::Array.new
      response = Concurrent::Hash.new
      threads = Concurrent::Array.new
      
      proposals_and_typers.each do |proposal_number, typer|
        threads << Thread.new do
          unless response.include? proposal_number or failed_proposals.include? proposal_number
            uf = nil
            con_pool.with do |con|
              uf = queryDatabaseUfOfProposal con, proposal_number
            end
            if uf
              response[proposal_number] = [uf, typer]
            else
              failed_proposals << proposal_number
            end
            mutex.synchronize {progress_keeper.progress += 1 unless progress_keeper.nil?}
          end
        end
      end
    ensure
      puts "Waiting processing of proposals"
      threads.each(&:join) unless threads.nil?
      puts "Threads finished"
      con_pool.shutdown { |con| con.close } unless con_pool.nil?
      puts "Connections closed"
    end
    return [response.collect {|i| i.flatten}, failed_proposals]
  end
  
  def recoverProposalNumbersAndStateOfProposals(file, bank, progress_keeper=nil)
    proposals_and_typers = acquireListOfProposals file, bank
    if proposals_and_typers.empty?
      raise "Nenhuma proposta localizada"
    else
      puts "Quantidade de propostas: #{proposals_and_typers.length}"
    end
    progress_keeper.total = proposals_and_typers.length unless progress_keeper.nil?
    findUfOfEachProposal proposals_and_typers, progress_keeper
  end
end

if __FILE__ == $0 ### script de teste ###
  include ProcessadorDePendencias
  
    abort "Passe apenas o caminho do arquivo excel como parametro" unless ARGV.length == 1
  recoverProposalNumbersAndStateOfProposal(ARGV[0], "teste")
end