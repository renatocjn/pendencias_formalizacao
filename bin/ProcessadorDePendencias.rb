module ProcessadorDePendencias
	require 'java'
	require 'rubygems'
	require 'jdbc/mysql'
    Jdbc::MySQL.load_driver
	java_import "com.mysql.jdbc.Driver"

	BANK_TO_SPREADSHEET_COLUMN_NUMBER_OF_PROPOSALS = {
		"teste" => 'E'
	}
	
	def acquireListOfProposals(file, bank)
		puts "acquireListOfProposals '#{file}', '#{bank}'"
		
		if file.end_with?("xls")
			require "roo-xls"
		else
			require "roo"
		end
		spreadsheet =  Roo::Spreadsheet.open(file)
		
		puts "column_number = '#{BANK_TO_SPREADSHEET_COLUMN_NUMBER_OF_PROPOSALS[bank]}'"
		column_number = BANK_TO_SPREADSHEET_COLUMN_NUMBER_OF_PROPOSALS[bank]
		
		spreadsheet.column(column_number)
	end
	
	def createDatabaseConnection
		puts "createDatabaseConnection"
		
		database_url = ""
		puts "database_url = '#{database_url}'"
		
		database_name = ""
		puts "database_name = '#{database_name}'"
		
		user_login = ""
		puts "user_login = '#{user_login}'"
		
		user_passwd = ""
		puts "user_passwd = '#{user_passwd}'"
		
		#con = Mysql.new(database_url, user_login, user_passwd, database_name)
		#begin
		#	con = java.sql.DriverManager.getConnection(
		#		"jdbc:mysql://#{database_url}/#{database_name}", 
		#		user_login, user_passwd
		#)
		#rescue
		#	abort "Failed to connect to database"
		#end
		#puts "Connection #{con.inspect}"
		#return con
		return nil
	end
	
	def findResponsableForEachProposal(proposals)
		con = createDatabaseConnection
		
		proposals.collect do |proposal_number|
			next unless proposal_number.is_a? Float or proposal_number.is_a? Integer
			
			uf = true
			raise "Failed to find UF for proposal #{proposal_number}" unless uf
			puts "proposal_number = '#{proposal_number}' | UF = '#{uf}'"
			[proposal_number, uf]
		end
	end
	
	def recoverProposalNumbersAndStateOfProposal(file, bank)
		proposals = acquireListOfProposals file, bank
		findResponsableForEachProposal proposals
	end
end

if __FILE__ == $0 ### script de teste ###
	include ProcessadorDePendencias
	
    abort "Passe apenas o caminho do arquivo excel como parametro" unless ARGV.length == 1
	recoverProposalNumbersAndStateOfProposal(ARGV[0], "teste")
end