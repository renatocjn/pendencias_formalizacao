#/usr/bin/env ruby

require 'fox16'
require 'Clipboard'
require 'logger'
require_relative "ProcessadorDePendencias"
include Fox

class ProcessadorDePendenciasGUI < FXMainWindow
  include ProcessadorDePendencias
  
  def initialize(app)
    Thread.abort_on_exception=true
    # Invoke base class initialize first
    super(app, "Processador de pendencias", opts: DECOR_ALL)
    
    @errorLog = Logger.new 'erros.log', 10
    @missingProposalsLog = Logger.new 'Propostas não encontradas.log', "daily"
    @missingProposalsLog.formatter = proc do |severity, datetime, progname, msg|
      " #{datetime.strftime("%d/%m/%Y %H:%M:%S")} | #{msg}\n"
    end
    @elements = Array.new 
    
    #Create a tooltip
    FXToolTip.new(self.getApp())
  
  
    ### Controls on top
    controls = FXVerticalFrame.new(self, LAYOUT_SIDE_TOP|LAYOUT_FILL_X,
      padLeft: 30, padRight: 30, padTop: 10, padBottom: 10)
  
    @elements << controlGroup = FXGroupBox.new(controls, "Selecione o banco sendo processado", GROUPBOX_TITLE_CENTER|FRAME_RIDGE|LAYOUT_SIDE_TOP|LAYOUT_FILL_X)
    @elements << optionsPopup = FXPopup.new(controlGroup)
    ["Detecção automática", "Itaú", "OLÉ", "HELP", "INTERMED EMPREST", "INTERMED CART", "Daycoval", "Centelem", "Bradesco", "CCB", "Bons", "Safra", "Sabemi", "PAN Consignado", "PAN Cartão", "Banrisul"].each { |opt| FXOption.new(optionsPopup, opt) }
    @bank_select = FXOptionMenu.new controlGroup, optionsPopup, opts: FRAME_THICK|FRAME_RAISED|ICON_BEFORE_TEXT|LAYOUT_FILL_X
    
    @elements << controlGroup = FXGroupBox.new(controls, "Selecione a planilha a ser processada", GROUPBOX_TITLE_CENTER|FRAME_RIDGE|LAYOUT_SIDE_TOP|LAYOUT_FILL_X)
    @progress_keeper = FXProgressBar.new(controlGroup, opts: PROGRESSBAR_NORMAL|PROGRESSBAR_HORIZONTAL|LAYOUT_FILL_X|LAYOUT_SIDE_BOTTOM|PROGRESSBAR_PERCENTAGE)
    @progress_keeper.barColor = FXRGB(0, 150, 0)
    @progress_keeper.textColor = FXRGB(0, 0, 0)
    @selectSpreadsheetBttn = FXButton.new(controlGroup, "Selecionar planilha", opts: LAYOUT_FILL_X|FRAME_THICK|FRAME_RIDGE|BUTTON_DEFAULT)
    setupSelectSpreadsheetBttn
  
    @elements << FXLabel.new(self, "contato: helpdesk@casebras.com.br", opts: JUSTIFY_CENTER_X|LAYOUT_SIDE_BOTTOM)
    @elements << FXHorizontalSeparator.new(self, LAYOUT_SIDE_BOTTOM|LAYOUT_FILL_X|SEPARATOR_GROOVE)
    
    
    ### Tables
    
    @elements << failedProposalsFrame = FXVerticalFrame.new(self, LAYOUT_SIDE_RIGHT|FRAME_NONE|LAYOUT_FILL_Y, padRight: 30, padLeft: 10)    
    @elements << box = FXGroupBox.new(failedProposalsFrame, "Propostas não encontradas", GROUPBOX_TITLE_CENTER|FRAME_RIDGE|LAYOUT_FILL_Y|LAYOUT_FILL_X, padding: 10)
    
    @failedProposalsTable = nil
    @elements << FXButton.new(box, "Copiar para área de transferência", opts: LAYOUT_SIDE_BOTTOM|FRAME_THICK|FRAME_RIDGE|LAYOUT_FILL_X).connect(SEL_COMMAND) do 
      copyContentsOfTable @failedProposalsTable
    end
    
    @failedProposalsTable = FXTable.new(box, opts: LAYOUT_CENTER_X|LAYOUT_CENTER_Y|LAYOUT_FILL_X|LAYOUT_FILL_Y|TABLE_COL_SIZABLE)
    @failedProposalsTable.rowHeaderWidth = 30
    @failedProposalsTable.visibleRows = 10
    @failedProposalsTable.visibleColumns = 1
    @failedProposalsTable.setTableSize 100, 1
    @failedProposalsTable.editable = false
   
    
    @elements << processedProposalsFrame = FXVerticalFrame.new(self, LAYOUT_SIDE_LEFT|FRAME_NONE|LAYOUT_FILL_Y|LAYOUT_FILL_X, padLeft: 30, padRight: 10)
    @elements << box = FXGroupBox.new(processedProposalsFrame, "Propostas encontradas", GROUPBOX_TITLE_CENTER|FRAME_RIDGE|LAYOUT_FILL_Y|LAYOUT_FILL_X, padding: 10)
    
    @processedProposalsTable = nil
    @elements << FXButton.new(box, "Copiar para área de transferência", opts: LAYOUT_SIDE_BOTTOM|FRAME_THICK|FRAME_RIDGE|LAYOUT_FILL_X).connect(SEL_COMMAND) do 
      copyContentsOfTable @processedProposalsTable
    end
    
    @processedProposalsTable = FXTable.new(box, opts: LAYOUT_CENTER_X|LAYOUT_CENTER_Y|LAYOUT_FILL_X|LAYOUT_FILL_Y|TABLE_COL_SIZABLE)
    @processedProposalsTable.rowHeaderWidth = 30
    @processedProposalsTable.visibleRows = 10
    @processedProposalsTable.visibleColumns = 3
    @processedProposalsTable.setTableSize 20, 3
    @processedProposalsTable.editable = false
  end
  
  def create
    super
    show(PLACEMENT_SCREEN)
  end
  
  def insertProcessedProposalsToTable (processedProposals, failedProposals)
    if processedProposals.empty?
      @processedProposalsTable.setTableSize 1,3
    else
      @processedProposalsTable.setTableSize(processedProposals.length, processedProposals.collect(&:length).max)
    end
    processedProposals.each_with_index do |values, i|
      values.each_with_index do |v, j| 
        @processedProposalsTable.setItemText(i,j, v.to_s.strip) 
      end
    end
    rowNumber = if failedProposals.empty? then 1 else failedProposals.length end
    @failedProposalsTable.setTableSize rowNumber, 1
    failedProposals.each_with_index {|p, i| @failedProposalsTable.setItemText(i,0, p.to_s.strip)}
  end
  
  def processSpreadsheet filename
    @selectSpreadsheetBttn.connect(SEL_COMMAND) do
      FXMessageBox::information self, MBOX_OK, "   Aguarde...", "Aguarde o processamento da planilha!"
    end
    @selectSpreadsheetBttn.text = "Aguarde..."
    @processThread = Thread.new(self) do |window|
      Thread::abort_on_exception = true
      begin
        puts "Starting main process thread"
        processedProposals, failedProposals = recoverProposalNumbersAndStateOfProposals(filename, @bank_select.current.to_s, @progress_keeper)
        puts "Done processing proposals, proposals found: #{processedProposals.length - 1}, failed proposals: #{failedProposals.length}"
        
        unless failedProposals.empty?
          failedProposalsMessage = "As seguintes propostas não puderam ser localizadas: " + failedProposals.join(", ")
          @missingProposalsLog.info failedProposalsMessage
        end
        
        insertProcessedProposalsToTable processedProposals, failedProposals
          
        doneMessage = "#{processedProposals.length - 1} propostas encontradas"
        doneMessage += " e #{failedProposals.length} propostas não puderam ser localizadas" unless failedProposals.empty?
        getApp.addChore {FXMessageBox::warning self, MBOX_OK, "   Algo deu errado...", doneMessage}
      rescue RuntimeError => err
        @errorLog.error err
        getApp.addChore {FXMessageBox::error self, MBOX_OK, "   Algo deu errado...", err.message}
      rescue Exception => exception
        @errorLog.error exception
        getApp.addChore {FXMessageBox::warning self, MBOX_OK, "   Algo deu errado...", "Algo muito errado aconteceu, favor entre em contato com a equipe de suporte!"}
      ensure
        setupSelectSpreadsheetBttn
        @progress_keeper.progress = @progress_keeper.total = 0
        puts "Main thread finished"
      end
    end
  end
  
  def setupSelectSpreadsheetBttn
    @selectSpreadsheetBttn.text = "Selecionar planilha"
    @selectSpreadsheetBttn.connect(SEL_COMMAND) do
      dialog = FXFileDialog.new(self, "Selecione a planilha de pendencias")  
      dialog.selectMode = SELECTFILE_EXISTING
      dialog.patternList = ["Excel Files (*.xls,*.xlsx)"]  
      if dialog.execute != 0
        processSpreadsheet dialog.filename
      end
    end
  end
  
  def copyContentsOfTable table
    excelFriendlyContent = String.new
    nRows = table.getNumRows
    nCols = table.getNumColumns
    nRows.times do |r|
      arr = Array.new
      nCols.times do |c|
        arr << table.getItemText(r,c)
      end
      excelFriendlyContent += arr.join("\t") + "\n"
    end
    Clipboard.copy excelFriendlyContent.strip
  end
end

app = FXApp.new "ProcessadorDePendencias", "Processamento de planilhas de pendencias"
ProcessadorDePendenciasGUI.new app
app.create
app.run