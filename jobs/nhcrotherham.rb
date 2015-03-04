require 'google_drive'
require 'roo'

#set up some info
sheetkey="1D8aS8eQO2XTitiQIpZkMFpR0vPx1P3SnxublT0vXSCY"
sheetfull="https://docs.google.com/spreadsheet/ccc?key=" + sheetkey   #use this to see the actual sheet
sheetresults = "#gid=1962399894"                 #append this to the URL for the sheet this program will use (pubmed_result)
results = sheetfull + sheetresults      #hey look I did it for you...

#puts "Please enter your Gmail address"
#GOOGLE_MAIL = gets
#puts "Please enter your password"
#GOOGLE_PASSWORD = gets

GOOGLE_MAIL = 'info@asnglobal.in'
GOOGLE_PASSWORD = 'asngc12398'

SCHEDULER.every '2m', :first_in => 0 do |job|
# :first_in sets how long it takes before the job is first run. In this case, it is run immediately

	# This passes the login credentials to Roo without requiring every user change system environment variables
	s = Roo::Google.new(sheetkey, user: GOOGLE_MAIL, password: GOOGLE_PASSWORD) #Loading spreadsheet :-)
	s.default_sheet = 'Management MI'

	send_event('clinic_details', { cliniclocation:s.cell('C',2), clinicdate:s.cell('C',3) })
	send_event('target', { value:s.cell('C',8), max:s.cell('C',4), title:"Clinic Target #{s.cell('C',4).to_i}" })
	send_event('total_leads', { value:s.cell('C',8) })
	send_event('taxis_booked', { value:s.cell('F',8) })
	send_event('letters_dispatched', { value:s.cell('C',12) })
	
	#Populate the agent stats
	agent_stats = Hash.new
	
	for i in 2..s.last_row('Agents MI').to_i - 1
		agent_stats[s.cell('A',i,'Agents MI')] = { label: s.cell('A',i,'Agents MI'), value: s.cell('B',i,'Agents MI').to_i }
	end
	send_event('agent_stats', { items: agent_stats.values })
	#agent_stats.each_value {|value| puts value}
	
end