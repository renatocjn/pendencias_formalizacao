
#!/usr/local/bin/jruby

include Java

#import java.awt.Insets
import javax.swing.JButton
import javax.swing.JFrame
import javax.swing.JPanel
import javax.swing.JTextArea
import javax.swing.JScrollPane
import javax.swing.JFileChooser
import javax.swing.SwingUtilities
import java.awt.BorderLayout
import java.awt.GridLayout
import java.awt.Dimension
import javax.swing.filechooser.FileNameExtensionFilter


class Example < JFrame
	require "Clipboard"
	require_relative "ProcessadorDePendencias"
	
	include ProcessadorDePendencias
	include java.awt.event.ActionListener
	
	attr_accessor :result, :filechooser, :openFileButton, :copyFileButton
	
	def initialize
		super "Pendências Formalização"
	end

    def initUI
        panel = JPanel.new
		panel.setLayout BorderLayout.new
        self.getContentPane.add panel
		
		buttonPanel = JPanel.new
		buttonPanel.setLayout GridLayout.new(0,1)
		panel.add buttonPanel, BorderLayout::PAGE_START
		
		userDir = java.lang.System.getProperty("user.home");
		self.filechooser = JFileChooser.new File.join(userDir, "desktop")
		self.filechooser.addChoosableFileFilter FileNameExtensionFilter.new("Arquivos Excel", "xls", "xlsx")
		self.filechooser.setAcceptAllFileFilterUsed false
		self.openFileButton = JButton.new "Escolher planilha"
		self.openFileButton.addActionListener self
		buttonPanel.add self.openFileButton
		
		self.copyFileButton = JButton.new "Copiar resultado"
		self.copyFileButton.addActionListener self
		buttonPanel.add self.copyFileButton
		
        self.result = JTextArea.new
		self.result.setEditable false
		#result.setMargin(Insets.new (5, 5, 5, 5) #Could not get to work!
		logScrollPane = JScrollPane.new self.result
		panel.add logScrollPane, BorderLayout::CENTER
		
        self.setDefaultCloseOperation JFrame::EXIT_ON_CLOSE
        #self.setSize 300, 180
		self.setMinimumSize Dimension.new(300, 180)
		#self.setResizable false
        self.setLocationRelativeTo nil
        self.setVisible true
    end
	
	def actionPerformed(event)
		if event.source == openFileButton
			return_value = self.filechooser.showOpenDialog(self)
			if return_value == JFileChooser::APPROVE_OPTION
				resultList = recoverProposalNumbersAndStateOfProposal filechooser.getSelectedFile().get_path(), "teste"
				self.result.setText('')
				puts resultList.inspect
				resultList.each do |r|
					self.result.append(Integer(r[0]).to_s)
					self.result.append("\t")
					self.result.append(r[1].to_s)
					self.result.append("\n")
				end
			end
			self.result.setCaretPosition(self.result.getDocument().getLength())
		elsif event.source == copyFileButton
			Clipboard.copy result.getText
		end
	end
end

Example.new.initUI