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
    @logger = Logger.new 'processador_de_pendencias.log', "weekly"
    @elements = Array.new 
    
    #Create a tooltip
    FXToolTip.new(self.getApp())
  
    #Controls on top
    controls = FXVerticalFrame.new(self, LAYOUT_SIDE_TOP|LAYOUT_FILL_X,
      padLeft: 40, padRight: 40, padTop: 10, padBottom: 10)
  
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
  
    @elements << FXLabel.new(self, "contato: helpdesk@casebras.com.br", opts: JUSTIFY_RIGHT|LAYOUT_SIDE_BOTTOM)
    @elements << FXHorizontalSeparator.new(self, LAYOUT_SIDE_BOTTOM|LAYOUT_FILL_X|SEPARATOR_GROOVE)
  
    #footer = FXVerticalFrame.new(self, LAYOUT_SIDE_BOTTOM|LAYOUT_FILL_X,
    #  padLeft: 40, padRight: 40, padTop: 10, padBottom: 10)
     
    #FXButton.new(footer, "Copiar para área de transferência", opts: LAYOUT_SIDE_BOTTOM|FRAME_THICK|FRAME_RIDGE|LAYOUT_FILL_X).connect(SEL_COMMAND) do 
    #  excelFriendlyContent = String.new
    #  @table.numRows.times { |r| excelFriendlyContent += "#{@table.getItemText(r,0)}\t#{@table.getItemText(r,1)}\t#{@table.getItemText(r,2)}\n" }
    #  Clipboard.copy excelFriendlyContent
    #end
    
    # Contents
    @elements << contents = FXHorizontalFrame.new(self,
      LAYOUT_CENTER_X|FRAME_NONE|LAYOUT_FILL_X|LAYOUT_FILL_Y, :padding => 10)
     
    @table = FXTable.new(contents, opts: LAYOUT_CENTER_X|LAYOUT_CENTER_Y|LAYOUT_FILL_X|LAYOUT_FILL_Y|TABLE_COL_SIZABLE)
    
    #@table.borderColor = FXRGB(255, 255, 255)
    @table.visibleRows = 10
    @table.visibleColumns = 3
    @table.setTableSize 10, 3
    @table.editable = false
    #@table.setColumnText 0, "Proposta"
    #@table.setColumnText 1, "UF"
  end
  
  def create
    super
    show(PLACEMENT_SCREEN)
  end
  
  def insertProcessedProposalsToTable (processedProposals)
    if processedProposals.empty?
      @table.setTableSize 1,3
    else
      @table.setTableSize processedProposals.length, processedProposals.collect(&:length).max
    end
    
    processedProposals.each_with_index do |(proposal, uf, *other_values), idx|
      @table.setItemText(idx,0, proposal.to_s)
      @table.setItemText(idx,1, uf.to_s)
      other_values.each_with_index { |v, i| @table.setItemText(idx,2+i, v.to_s) }
    end
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
          #failedProposalsMessage = "As seguintes propostas não puderam ser localizadas:\n\n" + failedProposals.each_slice(4).collect{|s| s.join("       ")}.join("\n")
          #window.showWarning failedProposalsMessage
          failedProposalsMessage = "As seguintes propostas não puderam ser localizadas: " + failedProposals.join(", ")
          @logger.info failedProposalsMessage
        end
        insertProcessedProposalsToTable processedProposals
      #rescue RuntimeError => err
        #window.showError err.message
      rescue Exception => exception
        @logger.error exception
        raise exception
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
  
  def showError msg
    FXMessageBox::error window, MBOX_OK, "   Algo deu errado...", msg
  end
  
  def showWarning msg
    FXMessageBox::warning window, MBOX_OK, "   Algo deu errado...", msg
  end
end

app = FXApp.new "ProcessadorDePendencias", "Processamento de planilhas de pendencias"
ProcessadorDePendenciasGUI.new app
app.create
app.run