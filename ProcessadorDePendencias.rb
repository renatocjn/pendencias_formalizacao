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
		database_url = ""
		database_name = ""
		user_login = ""
		user_passwd = ""
  require 'tiny_tds'
    
  BANK_TO_SPREADSHEET_COLUMN_NUMBER_OF_PROPOSALS = {
    "teste" => 'E',
    "help" => 'E',
    "ole" => 'H',
    "itau" => 'I'
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
    column_number = getSpreadSheetColumn file, bank
    
    
    #puts spreadsheet.column(column_number)
    #raise "teste"
    spreadsheet.column(column_number).reject {|v| v.nil?}.select {|v| v.is_a?(Integer) or v.is_a?(Float) or v.is_integer?}
  end
  
  def createDatabaseConnection
    
    begin
      TinyTds::Client.new username: user_login, password: user_passwd, host: database_url, database: database_name
    rescue TinyTds::Error => err
      raise "Falha ao se conectar ao banco"
    end 
  end
  
  def getSQL proposal
    "SELECT UE.SGL_UNIDADE_EMPRESA, UE.SGL_UNIDADE_FEDERACAO, UE.NOM_UNIDADE_EMPRESA, UE.NOM_FANTASIA
      FROM [CBDATA].[dbo].[PROPOSTA_EMPRESTIMO] AS PE
        INNER JOIN [CBDATA].[dbo].[UNIDADE_EMPRESA] AS UE ON UE.COD_UNIDADE_EMPRESA = PE.COD_UNIDADE_EMPRESA
      WHERE PE.NUM_PROPOSTA = '#{proposal}' OR PE.NUM_CONTRATO = '#{proposal}'"
  end
  
  def queryDatabaseForUfOfProposal con, proposal
    sql = getSQL proposal
    result = con.execute sql
    row = result.first
    if row.nil?
      nil
    else
      uf = row["SGL_UNIDADE_EMPRESA"]
      uf = uf.nil? ? row["SGL_UNIDADE_FEDERACAO"] : uf
      uf = uf.nil? ? row["NOM_UNIDADE_EMPRESA"] : uf
      uf = uf.nil? ? row["NOM_FANTASIA"] : uf
    end
  end
  
  def findUfOfEachProposal(proposals, progress_keeper=nil)
    con = createDatabaseConnection
    raise "Não foi possível acessar o banco de dados" unless con
    failed_proposals = Array.new
    response = proposals.collect do |proposal_number|
      uf = queryDatabaseForUfOfProposal con, proposal_number
      failed_proposals << proposal_number unless uf
      progress_keeper.progress += 1 unless progress_keeper.nil?
      [proposal_number, uf]
    end
    con.close
    return response, failed_proposals
  end
  
  def recoverProposalNumbersAndStateOfProposals(file, bank, progress_keeper=nil)
    proposals = acquireListOfProposals file, bank
    if proposals.empty? then raise "Nenhuma proposta localizada" end
    progress_keeper.total = proposals.length unless progress_keeper.nil?
    findUfOfEachProposal proposals, progress_keeper
  end
end

if __FILE__ == $0 ### script de teste ###
  include ProcessadorDePendencias
  
    abort "Passe apenas o caminho do arquivo excel como parametro" unless ARGV.length == 1
  recoverProposalNumbersAndStateOfProposal(ARGV[0], "teste")
end