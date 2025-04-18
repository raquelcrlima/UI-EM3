# Load libraries ----
library(shiny)
library(shinydashboard)
library(shinyjs)
library(readxl)
library(dplyr)
library(tidyr)
library(purrr)
library(FSA)
library(knitr)
library(rlang)
library(DT)
library(kableExtra)
library(openxlsx)

# Pre calculations ----
inputOptions <- data.frame(
  org = c("Company", "Company", "Company", "Company", "School", "School", "School", "School", "School", "School", "School",
          "School", "School", "School", "School", "School", "School", "School", "School", "School", "University", "University",
          "University", "University", "University", "University", "University", "University", "University", "University", "University", "University", "University",
          "University", "University", "University", "University", "University", "University", "University", "University", "University", "University", "University"
          ),
  previous = c("Yes", "Yes", "No", "No", "Yes", "Yes", "Yes", "Yes", "Yes", "Yes", "Yes",
               "Yes", "Yes", "Yes", "Yes", "Yes", "No", "No", "No", "No", "Yes", "Yes",
               "Yes", "Yes", "Yes", "Yes", "Yes", "Yes", "Yes", "Yes", "Yes", "Yes", "Yes",
               "Yes", "Yes", "Yes", "Yes", "Yes", "No", "No", "No", "No", "No", "No"
               ),
  curOld = c("No", "No", "Yes", "No", "No", "No", "No", "No", "No", "No", "No",
             "No", "No", "No", "No", "No", "Yes", "No", "No", "No", "No", "No",
             "No", "No", "No", "No", "No", "No", "No", "No", "No", "No", "No",
             "No", "No", "No", "No", "No", "Yes", "Yes", "Yes", "No", "No", "No"
             ),
  preOld = c("Yes", "No", NA, NA, "Yes", "Yes", "Yes", "No", "No", "No", "No",
             "No", "No", "No", "No", "No", NA, NA, NA, NA, "Yes", "Yes",
             "Yes", "Yes", "Yes", "Yes", "Yes", "Yes", "Yes", "No", "No", "No", "No",
             "No", "No", "No", "No", "No", NA, NA, NA, NA, NA, NA
             ),
  curAud = c("Employees", "Employees", "Employees", "Employees", "Students", "Employees", "Both", "Students", "Students", "Students", "Employees",
             "Employees", "Employees", "Both", "Both", "Both", "Students", "Students", "Employees", "Both", "Students", "Students",
             "Students", "Employees", "Employees", "Employees", "Both", "Both", "Both", "Students", "Students", "Students", "Employees",
             "Employees", "Employees", "Both", "Both", "Both", "Students", "Employees", "Both", "Students", "Employees", "Both"
             ),
  preAud = c("Employees", "Employees", NA, NA, "Students", "Students", "Students", "Students", "Employees", "Both", "Students",
             "Employees", "Both", "Students", "Employees", "Both", NA, NA, NA, NA, "Students", "Employees",
             "Both", "Students", "Employees", "Both", "Students", "Employees", "Both", "Students", "Employees", "Both", "Students",
             "Employees", "Both", "Students", "Employees", "Both", NA, NA, NA, NA, NA, NA
             ),
  comb = c(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11,
           2, 12, 13, 14, 15, 16, 17, 4, 18, 5, 19,
           20, 6, 1, 21, 7, 22, 23, 8, 9, 10, 11,
           2, 12, 13, 14, 15, 16, 3, 24, 17, 4, 18
           )
)

factorLookup <- data.frame(
  Response = c("11-13", "14-16", "17-19", "20-22", "23-29", "30-39", "40-49", "50-59", "60+",
               "Female", "Male", "Non-binary/other gender", "Prefer not to say",
               "Extremely responsible", "Very responsible", "Moderately responsible", "Slightly responsible", "Not at all responsible",
               "Strongly agree", "Agree", "Neutral", "Disagree", "Strongly disagree", "I don't have this information", "I don't know",
               "Always", "Often", "Sometimes", "Rarely", "Never",
               "Yes", "No",
               "More than once a year", "Yearly", "Less than once a year",
               "All", "Limited", "None"
               ),
  Order = c(1, 2, 3, 4, 5, 6, 7, 8, 9,
            1, 2, 3, 4,
            1, 2, 3, 4, 5,
            1, 2, 3, 4, 5, 6, 7,
            1, 2, 3, 4, 5,
            1, 2,
            1, 2, 3,
            1, 2, 3
            )
)

valueLookup <- data.frame(
  Response = c("Strongly agree", "Agree", "Neutral", "Disagree", "Strongly disagree", "I don't have this information", "I don't know",
               "Extremely responsible", "Very responsible", "Moderately responsible", "Slightly responsible", "Not at all responsible",
               "Always", "Often", "Sometimes", "Rarely", "Never", NA
               ),
  Value = c(5, 4, 3, 2, 1, NA, NA,
            5, 4, 3, 2, 1,
            5, 4, 3, 2, 1, NA
            )
)

catMeaning <- data.frame(
  Category = c("A", "B", "C", "D", "E"),
  Score = c(
    "Score between 4.2 (inclusive) and 5 (inclusive)",
    "Score between 3.4 (inclusive) and 4.2",
    "Score between 2.6 (inclusive) and 3.4",
    "Score between 1.8 (inclusive) and 2.6",
    "Score between 1 (inclusive) and 1.8"
  ),
  Definition = c(
    "Exemplary Maturity",
    "Advanced Maturity",
    "Intermediate Maturity",
    "Basic Maturity",
    "Initial Maturity"
  ),
  Description = c(
    "Best-in-class, to be used as an example",
    "Solid foundation, but there is room for minor improvements",
    "Has a foundation but needs improvement",
    "In early stages, needs substantial improvement",
    "Critical weakness, either absent or very weak, requires urgent attention"
  )
)

sigMeaning <- data.frame(
  Significance = c("ns", "*", "**", "***", "****"),
  Description = c(
    "The difference is not significant (it is unlikely that the difference is meaningful)",
    "The difference is significant (there is a 95% chance that the difference is meaningful)",
    "The difference is very significant (there is a 99% chance that the difference is meaningful)",
    "The difference is highly significant (there is a 99.9% chance that the difference is meaningful)",
    "The difference is extremely significant (there is a 99.99% chance that the difference is meaningful)"
  )
)

master <- data.frame(
  Code = c("Q002", "Q003", "Q189", "Q190", "Q191", "Q192", "Q029", "Q030", "Q031", "Q032", "Q033", "Q034", "Q035", "Q036", "Q037", "Q038", "Q039", "Q040", "Q041", "Q042", "Q044", "Q045", "Q046", "Q048", "Q049", "Q051", "Q052", "Q053", "Q054", "Q055", "Q056", "Q057", "Q058", "Q059", "Q060", "Q061", "Q062", "Q063", "Q064", "Q065", "Q066", "Q067", "Q068", "Q069", "Q070", "Q071", "Q075", "Q077", "Q078", "Q082", "Q105", "Q106", "Q107", "Q108", "Q109", "Q110", "Q111", "Q112", "Q113", "Q114", "Q116", "Q117", "Q119", "Q121", "Q122", "Q125", "Q127", "Q129", "Q130", "Q131", "Q139", "Q140", "Q141", "Q147", "Q153", "Q154", "Q155", "Q193", "Q194", "Q195"),
  q = c(NA, NA, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 4, 4, 4, 4, 4, 5, 5, 5, 5, 5, 5, 6, 6, 6, 6, 6, 6, 6, 6, 7, 7, 7, 7, 8, 8, 8, 9, 9, 9, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 6, 6, 6, 7, 7, 7),
  colOld = c("Please indicate your age", "Please indicate your gender", NA, NA, NA, NA, "Myself", "My family", "My organisation", "European Union", "National government (Ireland/Northern Ireland)", "Other governments (global)", "Local authorities", "Non-governmental organisations", "Businesses (companies)", "Scientists", "Adults, 18+ years old", "Teenagers, 12-17 years old", "I believe I know environmental issues well enough", "I believe I have enough skills to live a sustainable life", "I know how to save energy", "I believe small actions (such as turning off lights/appliances) make a difference in combating global environmental issues", "I consider environmental impacts when making choices (such as turning off lights/appliances)", "I am concerned about environmental issues for the protection of the natural environment", "Saving energy is important to me", "I save energy for environmental reasons", "I am motivated to save energy at home", "I am motivated to save energy in my organisation", "I discuss energy saving at home", "I discuss energy saving in my organisation", "I encourage my family to save energy", "I encourage my colleagues/fellow students to save energy", "I know how my organisation manages energy", "I know how to save energy in my organisation", "Energy conservation is a high priority activity in my organisation", "I am expected to try and save energy in my organisation", "I stay informed about my organisation's pro-environmental campaigns (such as posters, digital media, courses, activities, guidelines, etc.)", "It is easy to find a way to participate in pro-environmental activities in my organisation", "I actively participate in pro-environmental activities in my organisation", "I would engage more with energy savings in my organisation if I had more opportunity/control", "I often feel a comfortable temperature in my organisation", "I often feel a comfortable temperature at home", "I can provide feedback on the indoor/room temperature in my organisation", "I would be willing to provide feedback on the indoor/room temperature in my organisation", "Turn off the lights when I leave a room that won't be occupied", "Turn off the lights when there is sufficient daylight in a room", "Turn off electrical appliances/equipment when not in use", "Turn off the lights when I leave a room that won't be occupied 2", "Turn off the lights when there is sufficient daylight in a room 2", "Turn off electrical appliances/equipment when not in use 2", "There is a formal energy management system in my organisation", "My organisation follows energy management guidelines/standards (e.g., ISO 50001)", "My organisation follows environmental sustainability guidelines/standards (e.g., ISO 14001)", "My organisation has an internal guideline/plan on energy savings and other environmental ambitions (e.g., carbon emissions)", "Sustainability is part of the core activities/values of my organisation", "The organisation plans and adopts energy conservation measures on a regular basis", "The organisation audits its energy consumption on a regular basis; the energy management system is subject to continuous improvement", "I believe my organisation has enough resources to have a formal energy management system", "I believe my organisation has enough skills to have a formal energy management system", "I believe it makes sense for my organisation to have a formal energy management system", "My organisation has meters/sensors as part of its energy management", "The data from meters/sensors are automatically gathered", "There is a constant follow-up on building energy data that feeds into the energy management system", "Gather historical data", "Inform decision making", "Dissemination of information to users", "My organisation provides sustainability education/training/courses for its users (e.g., staff, students)", "Top managers are aware of the importance of energy conservation and support energy conservation initiatives", "Any user can effectively engage in the energy management of the organisation", "The users' input on sustainability matters and on energy management are an effective part of the decision-making process", "Thermal comfort", "Knowledge/attitudes regarding environmental or sustainability matters", "Sustainable practices in the organisation (e.g., use of lights, appliances, personal devices)", "How often does your organisation carry sustainability campaigns (e.g., turn it off)?", "Sustainability-related metrics and KPIs", "Results/insights from users' feedbacks and surveys", "Results/insights from sustainability campaigns", NA, NA, NA),
  Question = c("Please indicate your age", "Please indicate your gender", NA, NA, NA, NA, "Based on your opinion, please rank the levels of responsibility the following groups have in helping to combat global environmental issues", "Based on your opinion, please rank the levels of responsibility the following groups have in helping to combat global environmental issues", "Based on your opinion, please rank the levels of responsibility the following groups have in helping to combat global environmental issues", "Based on your opinion, please rank the levels of responsibility the following groups have in helping to combat global environmental issues", "Based on your opinion, please rank the levels of responsibility the following groups have in helping to combat global environmental issues", "Based on your opinion, please rank the levels of responsibility the following groups have in helping to combat global environmental issues", "Based on your opinion, please rank the levels of responsibility the following groups have in helping to combat global environmental issues", "Based on your opinion, please rank the levels of responsibility the following groups have in helping to combat global environmental issues", "Based on your opinion, please rank the levels of responsibility the following groups have in helping to combat global environmental issues", "Based on your opinion, please rank the levels of responsibility the following groups have in helping to combat global environmental issues", "Based on your opinion, please rank the levels of responsibility the following groups have in helping to combat global environmental issues", "Based on your opinion, please rank the levels of responsibility the following groups have in helping to combat global environmental issues", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "Please indicate how much you agree or disagree with the following statements", "How often do you perform the following activities AT HOME?", "How often do you perform the following activities AT HOME?", "How often do you perform the following activities AT HOME?", "How often do you perform the following activities IN YOUR ORGANISATION?", "How often do you perform the following activities IN YOUR ORGANISATION?", "How often do you perform the following activities IN YOUR ORGANISATION?", "Please rate the extent to which you agree/disagree with the following statements", "Please rate the extent to which you agree/disagree with the following statements", "Please rate the extent to which you agree/disagree with the following statements", "Please rate the extent to which you agree/disagree with the following statements", "Please rate the extent to which you agree/disagree with the following statements", "Please rate the extent to which you agree/disagree with the following statements", "Please rate the extent to which you agree/disagree with the following statements", "Please rate the extent to which you agree/disagree with the following statements", "Please rate the extent to which you agree/disagree with the following statements", "Please rate the extent to which you agree/disagree with the following statements", "Please rate the extent to which you agree/disagree with the following statements", "Please rate the extent to which you agree/disagree with the following statements", "Please rate the extent to which you agree/disagree with the following statements", "Please indicate whether your organisation have the following purposes for its meters/sensors", "Please indicate whether your organisation have the following purposes for its meters/sensors", "Please indicate whether your organisation have the following purposes for its meters/sensors", "Please rate the extent to which you agree/disagree with the following statements", "Please rate the extent to which you agree/disagree with the following statements", "Please rate the extent to which you agree/disagree with the following statements", "Please rate the extent to which you agree/disagree with the following statements", "How often does your organisation gather information/feedback from its users regarding the following topics?", "How often does your organisation gather information/feedback from its users regarding the following topics?", "How often does your organisation gather information/feedback from its users regarding the following topics?", "How often does your organisation carry sustainability campaigns (e.g., turn it off)?", "How much of the following data does your organisation disseminate with its users?", "How much of the following data does your organisation disseminate with its users?", "How much of the following data does your organisation disseminate with its users?", NA, NA, NA),
  Item = c(NA, NA, NA, NA, NA, NA, "Myself", "My family", "My organisation", "European Union", "National government (Ireland/Northern Ireland)", "Other governments (global)", "Local authorities", "Non-governmental organisations", "Businesses (companies)", "Scientists", "Adults, 18+ years old", "Teenagers, 12-17 years old", "I believe I know environmental issues well enough", "I believe I have enough skills to live a sustainable life", "I know how to save energy", "I believe small actions (such as turning off lights/appliances) make a difference in combating global environmental issues", "I consider environmental impacts when making choices (such as turning off lights/appliances)", "I am concerned about environmental issues for the protection of the natural environment", "Saving energy is important to me", "I save energy for environmental reasons", "I am motivated to save energy at home", "I am motivated to save energy in my organisation", "I discuss energy saving at home", "I discuss energy saving in my organisation", "I encourage my family to save energy", "I encourage my colleagues/fellow students to save energy", "I know how my organisation manages energy", "I know how to save energy in my organisation", "Energy conservation is a high priority activity in my organisation", "I am expected to try and save energy in my organisation", "I stay informed about my organisation's pro-environmental campaigns (such as posters, digital media, courses, activities, guidelines, etc.)", "It is easy to find a way to participate in pro-environmental activities in my organisation", "I actively participate in pro-environmental activities in my organisation", "I would engage more with energy savings in my organisation if I had more opportunity/control", "I often feel a comfortable temperature in my organisation", "I often feel a comfortable temperature at home", "I can provide feedback on the indoor/room temperature in my organisation", "I would be willing to provide feedback on the indoor/room temperature in my organisation", "Turn off the lights when I leave a room that won't be occupied", "Turn off the lights when there is sufficient daylight in a room", "Turn off electrical appliances/equipment when not in use", "Turn off the lights when I leave a room that won't be occupied", "Turn off the lights when there is sufficient daylight in a room", "Turn off electrical appliances/equipment when not in use", "There is a formal energy management system in my organisation", "My organisation follows energy management guidelines/standards (e.g., ISO 50001)", "My organisation follows environmental sustainability guidelines/standards (e.g., ISO 14001)", "My organisation has an internal guideline/plan on energy savings and other environmental ambitions (e.g., carbon emissions)", "Sustainability is part of the core activities/values of my organisation", "The organisation plans and adopts energy conservation measures on a regular basis", "The organisation audits its energy consumption on a regular basis; the energy management system is subject to continuous improvement", "I believe my organisation has enough resources to have a formal energy management system", "I believe my organisation has enough skills to have a formal energy management system", "I believe it makes sense for my organisation to have a formal energy management system", "My organisation has meters/sensors as part of its energy management", "The data from meters/sensors are automatically gathered", "There is a constant follow-up on building energy data that feeds into the energy management system", "Gather historical data", "Inform decision making", "Dissemination of information to users", "My organisation provides sustainability education/training/courses for its users (e.g., staff, students)", "Top managers are aware of the importance of energy conservation and support energy conservation initiatives", "Any user can effectively engage in the energy management of the organisation", "The users' input on sustainability matters and on energy management are an effective part of the decision-making process", "Thermal comfort", "Knowledge/attitudes regarding environmental or sustainability matters", "Sustainable practices in the organisation (e.g., use of lights, appliances, personal devices)", NA, "Sustainability-related metrics and KPIs", "Results/insights from users' feedbacks and surveys", "Results/insights from sustainability campaigns", NA, NA, NA),
  Ind = c("Please indicate your age", "Please indicate your gender", "I am taking or have taken part on environmental/sustainability education initiatives in my organisation (training, modules, courses, etc.)", "I am taking or have taken part on energy/energy management education initiatives in my organisation (training, modules, courses, etc.)", "My organisation contributed to my knowledge about global environmental issues (such as climate change, water pollution, biodiversity crises, etc.) and how to combat them", "My organisation contributed to my knowledge about energy and how to save it", "I am responsible for helping to combat global environmental issues", "My family is responsible for helping to combat global environmental issues", "My organisation is responsible for helping to combat global environmental issues", "The European Union is responsible for helping to combat global environmental issues", "The national government (Ireland/Northern Ireland) is responsible for helping to combat global environmental issues", "Other governments (global) are responsible for helping to combat global environmental issues", "Local authorities are responsible for helping to combat global environmental issues", "Non-governmental organisations are responsible for helping to combat global environmental issues", "Businesses (companies) are responsible for helping to combat global environmental issues", "Scientists are responsible for helping to combat global environmental issues", "Adults (18+ years old) are responsible for helping to combat global environmental issues", "Teenagers (12-17 years old) are responsible for helping to combat global environmental issues", "I know environmental issues well enough", "I have enough skills to live a sustainable life", "I know how to save energy", "Small actions (such as turning off lights/appliances) make a difference in combating global environmental issues", "I consider environmental impacts when making choices (such as turning off lights/appliances)", "I am concerned about environmental issues for the protection of the natural environment", "Saving energy is important to me", "I save energy for environmental reasons", "I am motivated to save energy at home", "I am motivated to save energy in my organisation", "I discuss energy saving at home", "I discuss energy saving in my organisation", "I encourage my family to save energy", "I encourage my colleagues/fellow students to save energy", "I know how my organisation manages energy", "I know how to save energy in my organisation", "Energy conservation is a high priority activity in my organisation", "I am expected to try and save energy in my organisation", "I stay informed about my organisation's pro-environmental campaigns (such as posters, digital media, courses, activities, guidelines, etc.)", "It is easy to find a way to participate in pro-environmental activities in my organisation", "I actively participate in pro-environmental activities in my organisation", "I would engage more with energy savings in my organisation if I had more opportunity/control", "I often feel a comfortable temperature in my organisation", "I often feel a comfortable temperature at home", "I can provide feedback on the indoor/room temperature in my organisation", "I would be willing to provide feedback on the indoor/room temperature in my organisation", "At home, I turn off the lights when I leave a room that won't be occupied", "At home, I turn off the lights when there is sufficient daylight in a room", "At home, I turn off electrical appliances/equipment when not in use", "In my organisation, I turn off the lights when I leave a room that won't be occupied", "In my organisation, I turn off the lights when there is sufficient daylight in a room", "In my organisation, I turn off electrical appliances/equipment when not in use", "There is a formal energy management system in my organisation", "My organisation follows energy management guidelines/standards (e.g., ISO 50001)", "My organisation follows environmental sustainability guidelines/standards (e.g., ISO 14001)", "My organisation has an internal guideline/plan on energy savings and other environmental ambitions (e.g., carbon emissions)", "Sustainability is part of the core activities/values of my organisation", "The organisation plans and adopts energy conservation measures on a regular basis", "The organisation audits its energy consumption on a regular basis; the energy management system is subject to continuous improvement", "My organisation has enough resources to have a formal energy management system", "My organisation has enough skills to have a formal energy management system", "It makes sense for my organisation to have a formal energy management system", "My organisation has meters/sensors as part of its energy management", "The data from meters/sensors are automatically gathered", "There is a constant follow-up on building energy data that feeds into the energy management system", "My organisation uses its sensors/meters to gather historical data", "My organisation uses its sensors/meters to inform decision making", "My organisation uses its sensors/meters to disseminate information to the users", "My organisation provides sustainability education/training/courses for its users", "Top managers are aware of the importance of energy conservation and support energy conservation initiatives", "Any user can effectively engage in the energy management of the organisation", "The users' input on sustainability matters and on energy management are an effective part of the decision-making process", "My organisation often (yearly) gathers information/feedback from its users regarding thermal comfort", "My organisation often (yearly) gathers information/feedback from its users regarding environmental/sustainability knowledge/attitudes", "My organisation often (yearly) gathers information/feedback from its users regarding sustainable practices in the organisation (e.g., use of lights, appliances, personal devices)", "My organisation often (yearly) carries sustainability campaigns (e.g., turn it off)", "My organisation disseminates to its users sustainability-related metrics and KPIs", "My organisation disseminates to its users results/insights from users' feedback and surveys", "My organisation disseminates to its users results/insights from sustainability campaigns", "In my organisation there is a direct feedback channel from users to energy management decision-makers (e.g., mobile app, feedback box/form, person to person, e-mail, phone, text message)", "In my organisation there are pro-environmental communication strategies (e.g., posters, digital screens/displays, prompts, newsletters)", "My organisation provides incentives to its users to engage in sustainability campaigns (e.g., financial incentive, promotion, bonus, competition/award, certificate)"),
  onFilter = c(1, 1, 2, 2, 2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2),
  audFilter = c(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2),
  rValue = c(NA, NA, 2, 2, 2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2),
  Model = c(NA, NA, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 1, 1, 1, 1, 1, 1, 1, NA, 1, NA, 1, NA, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, NA, 1, 1, NA, NA, NA, 2, 2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2),
  TopicCode = c(NA, NA, 6, 6, 6, 6, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 2, 1, 2, 2, 2, 3, 2, NA, 3, NA, 3, NA, 3, 1, 1, 5, 5, 3, 3, 3, 3, 4, NA, 7, 7, NA, NA, NA, 3, 3, 3, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 7, 6, 5, 3, 7, 7, 7, 7, 6, 7, 7, 7, 7, 7, 3),
  Topic = c(NA, NA, "Education", "Education", "Education", "Education", "Beliefs", "Beliefs", "Beliefs", "Beliefs", "Beliefs", "Beliefs", "Beliefs", "Beliefs", "Beliefs", "Beliefs", "Beliefs", "Beliefs", "Knowledge and Awareness", "Beliefs", "Knowledge and Awareness", "Beliefs", "Beliefs", "Beliefs", "Behaviours", "Beliefs", NA, "Behaviours", NA, "Behaviours", NA, "Behaviours", "Knowledge and Awareness", "Knowledge and Awareness", "Energy Management", "Energy Management", "Behaviours", "Behaviours", "Behaviours", "Behaviours", "Comfort", NA, "Feedback and Communication", "Feedback and Communication", NA, NA, NA, "Behaviours", "Behaviours", "Behaviours", "Energy Management", "Energy Management", "Energy Management", "Energy Management", "Energy Management", "Energy Management", "Energy Management", "Energy Management", "Energy Management", "Energy Management", "Energy Management", "Energy Management", "Energy Management", "Energy Management", "Energy Management", "Feedback and Communication", "Education", "Energy Management", "Behaviours", "Feedback and Communication", "Feedback and Communication", "Feedback and Communication", "Feedback and Communication", "Education", "Feedback and Communication", "Feedback and Communication", "Feedback and Communication", "Feedback and Communication", "Feedback and Communication", "Behaviours"),
  CritCode = c(NA, NA, 61, 62, 61, 62, 21, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 11, 21, 12, 21, 22, 22, 31, 22, NA, 31, NA, 32, NA, 32, 12, 12, 51, 51, 32, 32, 32, 32, 41, NA, 71, 71, NA, NA, NA, 33, 33, 33, 53, 53, 53, 53, 51, 53, 53, 52, 52, 52, 54, 54, 54, 54, 54, 72, 61, 51, 32, 71, 71, 71, 71, 61, 72, 72, 72, 71, 72, 32),
  Crit = c(NA, NA, "Sustainability Education", "Energy Management Education", "Sustainability Education", "Energy Management Education", "Self-Efficacy", "Perceived Broad Responsibility", "Perceived Broad Responsibility", "Perceived Broad Responsibility", "Perceived Broad Responsibility", "Perceived Broad Responsibility", "Perceived Broad Responsibility", "Perceived Broad Responsibility", "Perceived Broad Responsibility", "Perceived Broad Responsibility", "Perceived Broad Responsibility", "Perceived Broad Responsibility", "Environmental Awareness", "Self-Efficacy", "Energy Management Knowledge", "Self-Efficacy", "Environmental Concern", "Environmental Concern", "Motivation", "Environmental Concern", NA, "Motivation", NA, "Engagement", NA, "Engagement", "Energy Management Knowledge", "Energy Management Knowledge", "Values", "Values", "Engagement", "Engagement", "Engagement", "Engagement", "Thermal Comfort", NA, "Feedback", "Feedback", NA, NA, NA, "Practices", "Practices", "Practices", "Implementation", "Implementation", "Implementation", "Implementation", "Values", "Implementation", "Implementation", "Capabilities", "Capabilities", "Capabilities", "Monitoring ", "Monitoring ", "Monitoring ", "Monitoring ", "Monitoring ", "Communication", "Sustainability Education", "Values", "Engagement", "Feedback", "Feedback", "Feedback", "Feedback", "Sustainability Education", "Communication", "Communication", "Communication", "Feedback", "Communication", "Engagement"),
  IndCode = c(NA, NA, 6103, 6201, 6104, 6202, 2103, 2301, 2302, 2303, 2304, 2305, 2306, 2307, 2308, 2309, 2310, 2311, 1101, 2102, 1201, 2101, 2201, 2202, 3101, 2203, NA, 3102, NA, 3206, NA, 3207, 1202, 1203, 5103, 5104, 3203, 3204, 3205, 3201, 4101, NA, 7103, 7101, NA, NA, NA, 3301, 3302, 3303, 5301, 5302, 5303, 5304, 5101, 5305, 5306, 5201, 5202, 5203, 5401, 5402, 5403, 5404, 5405, 7204, 6101, 5102, 3202, 7102, 7104, 7105, 7106, 6102, 7201, 7202, 7203, 7107, 7205, 3208),
  tAna = c(NA, NA, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA),
  gAna = c("Age", "Gender", 2, 2, 2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA),
  hAna = c(NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, "I am motivated to save energy [at home vs. in my organisation]", "I am motivated to save energy [at home vs. in my organisation]", "I discuss energy saving [at home vs. in my organisation]", "I discuss energy saving [at home vs. in my organisation]", "I encourage my [family vs. colleagues/fellow students] to save energy", "I encourage my [family vs. colleagues/fellow students] to save energy", NA, NA, NA, NA, NA, NA, NA, NA, "I often feel a comfortable temperature [at home vs. in my organisation]", "I often feel a comfortable temperature [at home vs. in my organisation]", NA, NA, "[At home vs. In my organisation], I turn off the lights when I leave a room that won't be occupied", "[At home vs. In my organisation], I turn off the lights when there is sufficient daylight in a room", "[At home vs. In my organisation], I turn off electrical appliances/equipment when not in use", "[At home vs. In my organisation], I turn off the lights when I leave a room that won't be occupied", "[At home vs. In my organisation], I turn off the lights when there is sufficient daylight in a room", "[At home vs. In my organisation], I turn off electrical appliances/equipment when not in use", NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA),
  Setting = c(NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, "Home", "Organisation", "Home", "Organisation", "Home", "Organisation", NA, NA, NA, NA, NA, NA, NA, NA, "Organisation", "Home", NA, NA, "Home", "Home", "Home", "Organisation", "Organisation", "Organisation", NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA)
)

modelMeaning <- master %>% select(Topic, Crit, Ind, IndCode) %>% filter(!is.na(IndCode)) %>% arrange(IndCode) %>% select(-IndCode)
colnames(modelMeaning) <- c("Topics", "Criteria", "Indicators")

dmOld <- master %>% filter(onFilter == 1 & audFilter == 2) %>% select(Code, colOld)
dmNew <- master %>% filter(audFilter == 2) %>% select(Code, Ind)
userOld <- master %>% filter(onFilter == 1 & audFilter == 1) %>% select(Code, colOld)
userNew <- master %>% filter(audFilter == 1) %>% select(Code, Ind)

dfColNames <- c("Code", "Question")

colnames(dmOld) <- dfColNames
colnames(dmNew) <- dfColNames
colnames(userOld) <- dfColNames
colnames(userNew) <- dfColNames

comb <- list(list(dmNew, userNew, NULL, dmOld, userOld, NULL),
             list(dmNew, userNew, NULL, dmNew, userNew, NULL),
             list(dmOld, userOld, NULL, NULL, NULL, NULL),
             list(dmNew, userNew, NULL, NULL, NULL, NULL),
             list(dmNew, NULL, userNew, dmOld, NULL, userOld),
             list(dmNew, userNew, NULL, dmOld, NULL, userOld),
             list(dmNew, userNew, userNew, dmOld, NULL, userOld),
             list(dmNew, NULL, userNew, dmNew, NULL, userNew),
             list(dmNew, NULL, userNew, dmNew, userNew, NULL),
             list(dmNew, NULL, userNew, dmNew, userNew, userNew),
             list(dmNew, userNew, NULL, dmNew, NULL, userNew),
             list(dmNew, userNew, NULL, dmNew, userNew, userNew),
             list(dmNew, userNew, userNew, dmNew, NULL, userNew),
             list(dmNew, userNew, userNew, dmNew, userNew, NULL),
             list(dmNew, userNew, userNew, dmNew, userNew, userNew),
             list(dmOld, NULL, userOld, NULL, NULL, NULL),
             list(dmNew, NULL, userNew, NULL, NULL, NULL),
             list(dmNew, userNew, userNew, NULL, NULL, NULL),
             list(dmNew, NULL, userNew, dmOld, userOld, NULL),
             list(dmNew, NULL, userNew, dmOld, userOld, userOld),
             list(dmNew, userNew, NULL, dmOld, userOld, userOld),
             list(dmNew, userNew, userNew, dmOld, userOld, NULL),
             list(dmNew, userNew, userNew, dmOld, userOld, userOld),
             list(dmOld, userOld, userOld, NULL, NULL, NULL)
             )

for(i in 1:24) {
  names(comb[[i]]) <- c("curDm", "curEmp", "curStu", "preDm", "preEmp", "preStu")
}

# App UI ----
ui <- dashboardPage(
  dashboardHeader(title = "UI-EM3 App"),
  dashboardSidebar(
    sidebarMenu(id = "sidebarMenu",
                menuItem("Home", tabName = "home", icon = icon("home")),
                uiOutput("dynamicMenu1"),
                uiOutput("dynamicMenu2")
    )
  ),
  dashboardBody(
    useShinyjs(),
    tags$head(tags$style(HTML(".btn-enabled { background-color: #009f28; border-color: #009f28; color: white; }
    .btn-disabled { background-color: #d9534f; border-color: #d9534f; color: white; cursor: not-allowed; }"
                              )
                         )
              ),
    tabItems(
      tabItem(tabName = "home",
              h3("Home"),
              fluidPage(
                h4("About"),
                p("This app was derived from a PhD research at the University of Galway. It helps decision-makers in companies, universities, and schools identify key focus areas for improving energy management by integrating user perspectives."),
                h4("Terms used"),
                tags$ul(
                  tags$li(tags$b("UI-EM3: "), "User-Integrated Energy Management Maturity Model."),
                  tags$li(tags$b("Organisation: "), "Refers to universities, schools, or companies."),
                  tags$li(tags$b("Decision-makers: "), "Individuals responsible for energy management."),
                  tags$li(tags$b("Users: "), "Employees and/or students.")
                ),
                h4("Inputs"),
                tags$ul(
                  tags$li(tags$b("Use the provided survey templates: "), "Circulate the survey templates within your organisation as instructed. Do not modify, add, or remove any parts of the templates. Ensure they reach the appropriate audience (employees, students, or decision-makers)."),
                  tags$li(tags$b("Required surveys: "), "To proceed, circulate the decision-maker survey and at least one user survey (for employees or students)."),
                  tags$li(tags$b("Download and prepare the results: "), "After collecting responses, download the results in Excel format without making any changes to the files. Examples of the expected Excel files are provided below."),
                  tags$li(tags$b("Upload previous results (optional): "), "If you have previous results from these surveys, you can upload them as well to track changes over time."),
                  tags$li(tags$b("Start and upload files: "), "Click the “Start” button at the bottom of this page to go to the “Input Files” tab. There, select the relevant options and upload all required files. Once the uploads are complete, click “Continue” to view the results.")
                ),
                h4("Survey templates"),
                p("Please use the survey templates below. Do not modify, add, or remove any parts of the templates when circulating them."),
                tags$ul(
                  tags$li(tags$a(href = "https://forms.office.com/Pages/ShareFormPage.aspx?id=hrHjE0bEq0qcbZq5u3aBbJ-d3H8QHGBOhWKkyT-PCx1UQ1U1RERVSURXWUpSTFdUV1VIQlhBTEs4Vi4u&sharetoken=WqC5EDMLaet3WmDJyl5n", "Decision-makers")),
                  tags$li(tags$a(href = "https://forms.office.com/Pages/ShareFormPage.aspx?id=hrHjE0bEq0qcbZq5u3aBbJ-d3H8QHGBOhWKkyT-PCx1UMVkyVTJRSE9QRTJJQ0RXVVNWV0dNUTNHSi4u&sharetoken=rZZ7I4xrEFWQ1mHLVANg", "Employees")),
                  tags$li(tags$a(href = "https://forms.office.com/Pages/ShareFormPage.aspx?id=hrHjE0bEq0qcbZq5u3aBbJ-d3H8QHGBOhWKkyT-PCx1UQ1pQNkI1UDQ0VUpFT0RYQzhXOFRFMU4wVS4u&sharetoken=rZZ7I4xrEFWQ1mHLVANg", "University students")),
                  tags$li(tags$a(href = "https://forms.office.com/Pages/ShareFormPage.aspx?id=hrHjE0bEq0qcbZq5u3aBbJ-d3H8QHGBOhWKkyT-PCx1UQUpNVUdEVlhQRkpTM1JNQkNZT085T01XNy4u&sharetoken=rZZ7I4xrEFWQ1mHLVANg", "School students"))
                ),
                h4("Expected files (examples)"),
                p("Below are examples of how the input files (Excel sheets containing survey results) should look like."),
                tags$ul(
                  tags$li(downloadLink("downloadDm", "Decision-makers")),
                  tags$li(downloadLink("downloadEmp", "Employees")),
                  tags$li(downloadLink("downloadStuUni", "University students")),
                  tags$li(downloadLink("downloadStuSch", "School students"))
                ),
                h4("UI-EM3 overview"),
                p("The UI-EM3 is a model that assesses how well energy management is implemented and considers users’ perspectives. It uses survey data to measure various aspects of energy management and practices, broken down into indicators, criteria, and topics. The details of these components are provided in the table below."),
                uiOutput("modelDescription"),
                tags$style(HTML("table, th, td { border: 1px solid black; }")),
                tags$ul(
                  tags$li(tags$b("Scoring: "), "Survey responses are rated from 1 to 5. Each indicator’s score is the average of its survey responses. Similarly, the scores for criteria and topics are averaged from their components."),
                  tags$li(tags$b("Maturity Categories: "), "Scores are grouped into categories to help understand the maturity level of energy management, as shown in the table below."),
                ),
                uiOutput("catDescription"),
                h4("Interpretation examples"),
                p("Examples of how the maturity of indicators, criteria and topics can be understood, based on their category, are given below."),
                tags$ul(
                  tags$li("If the indicator “I turn off the lights when I leave a room” is rated ", tags$b("A"), ", it means this behaviour is consistently practiced and should be used as an example."),
                  tags$li("If the indicator “I turn off the lights when I leave a room” is rated ", tags$b("E"), ", it means this behaviour is rarely or never practiced, and immediate action is needed."),
                  tags$li("If the criterium “Environmental Awareness” is rated", tags$b("B"), ",  it means the users have strong awareness but could improve in specific areas."),
                  tags$li("If the criterium “Environmental Awareness” is rated", tags$b("D"), ", it means awareness is limited, and significant effort is needed to improve."),
                  tags$li("If the topic “Knowledge and Awareness” is rated", tags$b("C"), ", it means the users have moderate knowledge and awareness but needs to strengthen specific criteria."),
                  tags$li("If the topic “Knowledge and Awareness” is rated", tags$b("E"), ", it means they lack basic knowledge and awareness, and this area requires urgent attention."),
                  tags$li("If an item has missing score/category (", tags$b("NA"), "), this item should be further investigated to assess whether there are weakenesses that require immediate action.")
                ),
                h4("Comparisons"),
                p("In addition to UI-EM3, the app performs statistical comparisons to provide deeper insights:"),
                tags$ul(
                  tags$li(tags$b("Setting: "), "Compare practices at home vs. in the organisation for spill-over analysis."),
                  tags$li(tags$b("Audience: "), "Compare students vs. employees (where applicable) to identify similarities and differences in approaches that target these groups."),
                  tags$li(tags$b("Time: "), "Compare previous vs. current results (where applicable) to track changes over time.")
                ),
                p("The significance of differences between these groups is provided in the table below."),
                uiOutput("sigDescription"),
                h4("Outputs"),
                p("After uploading all required files and clicking “Continue”, three tabs will appear:"),
                tags$ul(
                  tags$li(tags$b("Highlights: "), "Provides a roadmap for improvement, highlighting key focus areas to target."),
                  tags$li(tags$b("Results: "), "Displays an overview of survey results, UI-EM3 outcomes, and comparisons across settings, audience, and time.")
                ),
                h4("More information"),
                p("If you have any questions, please contact Raquel Lima (raquel.lima@universityofgalway.ie) or refer to the associated research."),
                hr(),
                actionButton("showMenu", "Start", class = "btn-enabled"),
                hr()
              )
      ),
      tabItem(tabName = "dataUpload",
              h3("Data Upload"),
              h4("Please upload all the asked files and click the button 'Start' to proceed. For more information, please consult the Home tab."),
              br(),
              fluidRow(
                box(radioButtons("org", "Please select your organisation", choices = c("University", "School", "Company"), selected = "University", inline = FALSE),
                    radioButtons("previous", "Do you want to track changes over time?", choices = c("Yes", "No"), selected = "Yes", inline = FALSE),
                    radioButtons("curOld", "Were the CURRENT surveys circulated before 2024?", choices = c("Yes", "No"), selected = "Yes", inline = FALSE),
                    radioButtons("preOld", "Were the PREVIOUS surveys circulated before 2024?", choices = c("Yes", "No"), selected = "Yes", inline = FALSE),
                    radioButtons("curAud", "What users were targeted by the CURRENT surveys?", choices = c("Students", "Employees", "Both"), selected = "Both", inline = FALSE),
                    radioButtons("preAud", "What users were targeted by the PREVIOUS surveys?", choices = c("Students", "Employees", "Both"), selected = "Both", inline = FALSE),
                    width = 4
                ),
                box(id = "fileInputs",
                    fileInput("curDm", "Please upload the responses to the CURRENT survey circulated to DECISION-MAKERS", multiple = FALSE, accept = c(".xlsx")),
                    uiOutput("curDmError"),
                    fileInput("curEmp", "Please upload the responses to the CURRENT survey circulated to EMPLOYEES", multiple = FALSE, accept = c(".xlsx")),
                    uiOutput("curEmpError"),
                    fileInput("curStu", "Please upload the responses to the CURRENT survey circulated to STUDENTS", multiple = FALSE, accept = c(".xlsx")),
                    uiOutput("curStuError"),
                    fileInput("preDm", "Please upload the responses to the PREVIOUS survey circulated to DECISION-MAKERS", multiple = FALSE, accept = c(".xlsx")),
                    uiOutput("preDmError"),
                    fileInput("preEmp", "Please upload the responses to the PREVIOUS survey circulated to EMPLOYEES", multiple = FALSE, accept = c(".xlsx")),
                    uiOutput("preEmpError"),
                    fileInput("preStu", "Please upload the responses to the PREVIOUS survey circulated to STUDENTS", multiple = FALSE, accept = c(".xlsx")),
                    uiOutput("preStuError"),
                    actionButton("submit", "Submit", class = "btn-enabled"),
                    width = 8
                )
              )
      ),
      tabItem(tabName = "results",
              h3("Overview"),
              fluidRow(
                tabBox(id = "overviewResults", width = 12,
                  tabPanel(id = "overviewResultsDem", title = "Demographics",
                           uiOutput("selectOverviewDem"),
                           uiOutput("overviewDem"),
                           p("Note: 'N' is the total number of responses. For 'Audience', it considers only the responses to the current surveys. For 'Age' and 'Gender', it considers only the responses to the current user surveys (students and/or employees).")
                           ),
                  tabPanel(id = "overviewResultsDm", title = "Decision-makers",
                           uiOutput("selectOverviewDm"),
                           uiOutput("overviewDm"),
                           p("Note: 'N' is the total number of responses to the current decision-maker survey.")
                           ),
                  tabPanel(id = "overviewResultsUser", title = "Users",
                           uiOutput("selectOverviewUser"),
                           uiOutput("overviewUser"),
                           p("Note: 'N' is the total number of responses to the current user surveys (students and/or employees).")
                           )
                )
              ),
              h3("Comparisons"),
              fluidRow(
                tabBox(id = "overviewComp", width = 12,
                  tabPanel(id = "overviewCompSetting", title = "Setting",
                           uiOutput("hAna")
                  ),
                  tabPanel(id = "overviewCompAudience", title = "Audience",
                           uiOutput("gAna")
                  ),
                  tabPanel(id = "overviewCompTime", title = "Time",
                           uiOutput("tAna")
                  )
                )
              ),
              h3("Model results"),
              fluidRow(
                box(
                  uiOutput("modelResults"),
                  width = 12
                )
              )
      ),
      tabItem(tabName = "highlights",
              h3("Highlights"),
              fluidRow(
                tabBox(id = "highlightModel", width = 12,
                       tabPanel(id = "highlightCrit", title = "Criteria",
                                uiOutput("critRes")
                       ),
                       tabPanel(id = "highlightInd", title = "Indicators",
                                uiOutput("indRes")
                       ),
                       tabPanel(id = "highlightTopic", title = "Topics",
                                uiOutput("topicRes")
                       )
                )
              ),
              h3("Next steps"),
              fluidRow(
                box(
                  p(HTML("The <b>topics</b> show the key themes evaluated in the model, to have a clearer vision on <b>what is being measured</b>.<br/>
                  The <b>criteria</b> are the <b>main instrument to inform decision-making</b>, as they evaluate more specific themes under the topics, and are informed by the indicators.<br/>
                  The <b>indicators</b> can be <b>measured through questionnaires</b>.<br/><br/>
                  To improve energy management by integrating user perspectives, please <b>focus on the areas showed at the top of the CRITERIA table</b> above.<br/><br/>
                  As lack of data in the questionnaires are not accounted for when calculating the scores, <b>items with no score (NA) are placed at the top of the tables above</b>, 
                  as they are as important or more important than lower scores and <b>should be targeted first</b>.<br/>
                  As missing data can interfere with the ratings, please consult the INDICATORS table above to pinpoint specific areas that need improvement.<br/><br/>
                  Also, make sure to select respondents to the decision-making survey that have as many information asked as possible.<br/>
                  You can also consult the comparisons across setting, adience and time to get further insights.<br/>
                  Informed by the results of this tool, <b>it is recommended that the organisations conducts follow-up interviews</b> with both decision-makers and users to validate the responses, 
                  and get additional/more detailed insights.<br/><br/>
                  For more information and recommendations, please consult the 'Home' tab and refer to the associated research.")),
                  width = 12
                )
              ),
              h3("Downloads"),
              fluidRow(
                box(
                  downloadButton("downloadDem", "Download Demographics' Overview"),
                  br(), br(),
                  downloadButton("downloadDM", "Download Decision-makers' Overview"),
                  br(), br(),
                  downloadButton("downloadUser", "Download Users' Overview"),
                  br(), br(),
                  downloadButton("downloadComparison", "Download Comparisons"),
                  br(), br(),
                  downloadButton("downloadModel", "Download Model Results"),
                  width = 12
                )
              )
      )
    )
  )
)

# App server ----
server <- function(input, output, session) {
  
  rv <- reactiveValues(showMenu = FALSE,
                       showResult = FALSE)

  reactiveVals <- reactiveValues(
    preOld = NULL,
    preAud = NULL,
    fileCurDm = NULL,
    fileCurEmp = NULL,
    fileCurStu = NULL,
    filePreDm = NULL,
    filePreEmp = NULL,
    filePreStu = NULL,
    selectedFile = NULL,
    comparison = c(FALSE, FALSE, FALSE, FALSE, FALSE, FALSE),
    results = list(),
    dataWide = NULL,
    dataLong = NULL,
    dataFactor = NULL,
    dataValue = NULL,
    overviewSur = NULL,
    overviewAud = NULL,
    overviewAge = NULL,
    overviewGen = NULL,
    overviewQ = NULL,
    tRes1 = NULL,
    tRes2 = NULL,
    tRes3 = NULL,
    gRes1 = NULL,
    gRes2 = NULL,
    gRes3 = NULL,
    hRes1 = NULL,
    hRes2 = NULL,
    hRes3 = NULL,
    critRes = NULL,
    indRes = NULL,
    topicRes = NULL,
    choices = NULL,
    downloadDem = NULL,
    comparisons = NULL,
    modelRes = NULL,
    modelResLevel = NULL,
    modelResPrint = NULL
  )
  
  resetFileInputs <- function() {
    reset("curDm")
    reset("curEmp")
    reset("curStu")
    reset("preDm")
    reset("preEmp")
    reset("preStu")
    reactiveVals$fileCurDm <- NULL
    reactiveVals$fileCurEmp <- NULL
    reactiveVals$fileCurStu <- NULL
    reactiveVals$filePreDm <- NULL
    reactiveVals$filePreEmp <- NULL
    reactiveVals$filePreStu <- NULL
    reactiveVals$dataWide <- NULL
    reactiveVals$dataLong <- NULL
    reactiveVals$dataFactor <- NULL
    reactiveVals$dataValue <- NULL
    reactiveVals$overviewSur <- NULL
    reactiveVals$overviewAud <- NULL
    reactiveVals$overviewAge <- NULL
    reactiveVals$overviewGen <- NULL
    reactiveVals$overviewQ <- NULL
    reactiveVals$tRes1 <- NULL
    reactiveVals$tRes2 <- NULL
    reactiveVals$tRes3 <- NULL
    reactiveVals$gRes1 <- NULL
    reactiveVals$gRes2 <- NULL
    reactiveVals$gRes3 <- NULL
    reactiveVals$hRes1 <- NULL
    reactiveVals$hRes2 <- NULL
    reactiveVals$hRes3 <- NULL
    reactiveVals$critRes <- NULL
    reactiveVals$indRes <- NULL
    reactiveVals$topicRes <- NULL
    reactiveVals$choices <- NULL
    reactiveVals$downloadDem <- NULL
    reactiveVals$comparisons <- NULL
    reactiveVals$modelRes <- NULL
    reactiveVals$modelResLevel <- NULL
    reactiveVals$modelResPrint <- NULL
  }
  
  readExcelFile <- function(file) {
    if(is.null(file)) {
      return(NULL)
    }
    as.data.frame(read_excel(file$datapath))
  }
  
  acceptFile <- function(file) {
    if(is.null(file) || !is.data.frame(file) || ncol(file) <= 5) {
      return(NULL)
    } else if(ncol(file) > 5) {
      data <- file %>% select(-(1:5))
      return(data)
    }
  }
  
  errorSheet <- function() {
    return(div(style = "color: red;", "Error: The uploaded file is not valid. For more information, please consult the Home tab."))
  }
  
  overviewTable <- function(table){
    table1 <- table %>% summarise(count = n(), .groups = 'drop') %>% mutate(total = sum(count)) %>% ungroup() %>%
      mutate(percent = (count/total)*100, Count = paste(count, sprintf("(%.1f%%)", percent), sep = " "))
    N <- table1$total[1]
    table1 <- table1 %>% rename(!!paste0("Count (N = ", N, ")") := Count)
    return(table1)
  }
  
  wilcoxP <- function(table, col2){
    wilcox <- wilcox.test(Value ~ get(col2), data = table)
    pValue <- wilcox$p.value
    return(pValue)
  }
  
  significance <- function(x){
    case_when(x < 0.0001 ~ "****",
              x < 0.001 ~ "***",
              x < 0.01 ~ "**",
              x < 0.05 ~ "*",
              x >= 0.05 ~ "ns",
              TRUE ~ NA)
  }
  
  blankVector <- function(x){
    if(length(x) == 0){
      x <- "No results to display"
    }
    return(x)
  }
  
  dataModel <- function(table, survey, old){
    if(old == "No"){
      data <- table %>% filter(Survey == survey & !is.na(Model))
    } else if(old == "Yes") {
      data <- table %>% filter(Survey == survey & Model == 1)
    }
    data <- data %>% arrange(IndCode) %>% select(TopicCode, Topic, CritCode, Crit, IndCode, Ind, Value)
    return(data)
  }
  
  category <- function(x){
    case_when(x >= 4.2 ~ "A", x >= 3.4 ~ "B", x >= 2.6 ~ "C", x >= 1.8 ~ "D", x >= 1 ~ "E", TRUE ~ NA)
  }
  
  changeLabel <- function(change){
    if(is.na(change)){
      return("No previous data")
    } else if(change > 0){
      return(paste0("Increased (+", sprintf("%.1f%%", change), ")"))
    } else if(change < 0){
      return(paste0("Decreased (", sprintf("%.1f%%", change), ")"))
    } else {
      return("No change")
    }
  }
  
  modelCalc <- function(table, cur, part, colName){
    if(part == "Topic"){
      data <- table %>% group_by(TopicCode, Topic)
    } else if(part == "Crit"){
      data <- table %>% group_by(TopicCode, Topic, CritCode, Crit)
    } else {
      data <- table %>% group_by(TopicCode, Topic, CritCode, Crit, IndCode, Ind)
    }
    if(cur == "Current"){
      data <- data %>% summarise(Score = round(mean(!!sym(colName), na.rm = TRUE), digits = 2), Cat = category(Score), .groups = 'drop')
    } else {
      data <- data %>% summarise(Score = round(mean(!!sym(colName), na.rm = TRUE), digits = 2), Cat = category(Score), .groups = 'drop')
    }
    data <- data %>% mutate(Score = ifelse(is.nan(Score), NA, Score))
    return(data)
  }
  
  percentChange <- function(table, part){
    newCol <- paste0(part, "Change")
    curCol <- paste0(part, "Score")
    preCol <- paste0(part, "ScorePre")
    data <- table %>% mutate(!!newCol := ifelse(is.na(!!sym(curCol)) | is.na(!!sym(preCol)), NA, ((!!sym(curCol) - !!sym(preCol)) / !!sym(preCol)) * 100))
  }
  
  modelTables <- function(table, part){
    score <- paste0(part, "Score")
    cat <- paste0(part, "Cat")
    change <- paste0(part, "ChangeLabel")
    data <- table %>% select(any_of(c(part, score, cat, change))) %>% unique() %>% arrange(!!sym(score))
    return(data)
  }
  
  renameColumns <- function(table){
    oldNewNames <- c("TopicScore" = "Score", "TopicCat" = "Category", "TopicChangeLabel" = "Change",
                     "Crit" = "Criterium", "CritScore" = "Score", "CritCat" = "Category", "CritChangeLabel" = "Change",
                     "Ind" = "Indicator", "IndScore" = "Score", "IndCat" = "Category", "IndChangeLabel" = "Change")
    existingColumns <- intersect(names(table), names(oldNewNames))
    data <- table %>% rename_with(~ oldNewNames[.x], all_of(existingColumns))
    return(data)
  }
  
  compareCat <- function(cur, pre) {
    levels <- c("A", "B", "C", "D", "E")
    curIndex <- match(cur, levels)
    preIndex <- match(pre, levels)
    if (is.na(curIndex) || is.na(preIndex)) {
      return("No previous data")
    } else if (curIndex < preIndex) {
      return(paste0("Increased (from ", pre, ")"))
    } else if (curIndex > preIndex) {
      return(paste0("Decreased (from ", pre, ")"))
    } else {
      return("No change")
    }
  }
  
  writeListExcel <- function(list, file) {
    req(list)
    wb <- createWorkbook()
    for (name in names(list)) {
      addWorksheet(wb, name)
      writeData(wb, name, list[[name]])
    }
    saveWorkbook(wb, file, overwrite = TRUE)
  }
  
  writeListCaption <- function(list, file) {
    wb <- createWorkbook()
    for (i in seq_along(list)) {
      sheetName <- paste0("Sheet", i)
      addWorksheet(wb, sheetName)
      longName <- names(list)[i]
      writeData(wb, sheetName, longName, startCol = 1, startRow = 1)
      writeData(wb, sheetName, list[[i]], startCol = 1, startRow = 3)
    }
    saveWorkbook(wb, file, overwrite = TRUE)
  }
  
  writeDataframeExcel <- function(file) {
    req(reactiveVals$critRes)
    req(reactiveVals$indRes)
    req(reactiveVals$topicRes)
    req(reactiveVals$modelRes)
    wb <- createWorkbook()
    addWorksheet(wb, "Highlights_Indicators")
    writeData(wb, "Highlights_Indicators", reactiveVals$indRes)
    addWorksheet(wb, "Highlights_Criteria")
    writeData(wb, "Highlights_Criteria", reactiveVals$critRes)
    addWorksheet(wb, "Highlights_Topics")
    writeData(wb, "Highlights_Topics", reactiveVals$topicRes)
    addWorksheet(wb, "Model results")
    writeData(wb, "Model results", reactiveVals$modelResPrint)
    saveWorkbook(wb, file, overwrite = TRUE)
  }

  observe({
    if(input$org == "Company") {
      if(input$previous == "Yes") {
        updateRadioButtons(session, "curOld", selected = "No")
        hide("curOld")
        show("preOld")
        reactiveVals$preOld <- input$preOld
        updateRadioButtons(session, "curAud", selected = "Employees")
        hide("curAud")
        updateRadioButtons(session, "preAud", selected = "Employees")
        hide("preAud")
        reactiveVals$preAud <- input$preAud
      } else if(input$previous == "No") {
        show("curOld")
        reactiveVals$preOld <- NULL
        hide("preOld")
        updateRadioButtons(session, "curAud", selected = "Employees")
        hide("curAud")
        reactiveVals$preAud <- NULL
        hide("preAud")
      }
    } else if(input$org == "School") {
      if(input$previous == "Yes") {
        updateRadioButtons(session, "curOld", selected = "No")
        hide("curOld")
        show("preOld")
        reactiveVals$preOld <- input$preOld
        show("curAud")
        if(input$preOld == "Yes") {
          updateRadioButtons(session, "preAud", selected = "Students")
          hide("preAud")
          reactiveVals$preAud <- input$preAud
        } else if(input$preOld == "No") {
          show("preAud")
        }
      } else if(input$previous == "No") {
        show("curOld")
        reactiveVals$preOld <- NULL
        hide("preOld")
        if(input$curOld == "Yes") {
          updateRadioButtons(session, "curAud", selected = "Students")
          hide("curAud")
        } else if(input$curOld == "No") {
          show("curAud")
        }
        reactiveVals$preAud <- NULL
        hide("preAud")
      }
    } else if(input$org == "University") {
      if(input$previous == "Yes") {
        updateRadioButtons(session, "curOld", selected = "No")
        hide("curOld")
        show("preOld")
        reactiveVals$preOld <- input$preOld
        show("curAud")
        show("preAud")
        reactiveVals$preAud <- input$preAud
      } else if(input$previous == "No") {
        show("curOld")
        reactiveVals$preOld <- NULL
        hide("preOld")
        show("curAud")
        reactiveVals$preAud <- NULL
        hide("preAud")
      }
    }
  })
  
  observe({
    if(input$curAud == "Both") {
      show("curEmp")
      show("curStu")
    } else if(input$curAud == "Students") {
      reset("curEmp")
      reactiveVals$fileCurEmp <- NULL
      hide("curEmp")
      show("curStu")
    } else if(input$curAud == "Employees") {
      show("curEmp")
      reset("curStu")
      reactiveVals$fileCurStu <- NULL
      hide("curStu")
    }
  })
  
  observe({
    if(input$previous == "No" || is.null(input$preAud)) {
      reset("preDm")
      reactiveVals$filePreDm <- NULL
      hide("preDm")
      reset("preEmp")
      reactiveVals$filePreEmp <- NULL
      hide("preEmp")
      reset("preStu")
      reactiveVals$filePreStu <- NULL
      hide("preStu")
    } else if(!is.null(input$preAud) && input$preAud == "Both") {
      show("preDm")
      show("preEmp")
      show("preStu")
    } else if(!is.null(input$preAud) && input$preAud == "Students") {
      show("preDm")
      reset("preEmp")
      reactiveVals$filePreEmp <- NULL
      hide("preEmp")
      show("preStu")
    } else if(!is.null(input$preAud) && input$preAud == "Employees") {
      show("preDm")
      show("preEmp")
      reset("preStu")
      reactiveVals$filePreStu <- NULL
      hide("preStu")
    }
  })
  
  observeEvent(c(input$org, input$previous, input$curOld, input$curAud, input$preOld, input$preAud), {
    resetFileInputs()
    if(input$previous == "Yes") {
      selectedComb <- inputOptions %>% filter(org == input$org, previous == input$previous, curOld == input$curOld, curAud == input$curAud, preOld == input$preOld, preAud == input$preAud) %>% pull(comb)
    } else if(input$previous == "No") {
      selectedComb <- inputOptions %>% filter(org == input$org, previous == input$previous, curOld == input$curOld, curAud == input$curAud) %>% pull(comb)
    }
    if (length(selectedComb) > 0) {
      reactiveVals$selectedFile <- selectedComb
    } else {
      reactiveVals$selectedFile <- NULL
    }
  })
  
  observeEvent(input$curDm, {
    reactiveVals$fileCurDm <- acceptFile(readExcelFile(input$curDm))
  })
  
  observeEvent(input$curEmp, {
    reactiveVals$fileCurEmp <- acceptFile(readExcelFile(input$curEmp))
  })
  
  observeEvent(input$curStu, {
    reactiveVals$fileCurStu <- acceptFile(readExcelFile(input$curStu))
  })
  
  observeEvent(input$preDm, {
    reactiveVals$filePreDm <- acceptFile(readExcelFile(input$preDm))
  })
  
  observeEvent(input$preEmp, {
    reactiveVals$filePreEmp <- acceptFile(readExcelFile(input$preEmp))
  })
  
  observeEvent(input$preStu, {
    reactiveVals$filePreStu <- acceptFile(readExcelFile(input$preStu))
  })
  
  observe({
    files <- list(
      curDm = reactiveVals$fileCurDm,
      curEmp = reactiveVals$fileCurEmp,
      curStu = reactiveVals$fileCurStu,
      preDm = reactiveVals$filePreDm,
      preEmp = reactiveVals$filePreEmp,
      preStu = reactiveVals$filePreStu
    )
    
    selectedInput <- reactiveVals$selectedFile
    
    if (!is.null(selectedInput)) {
      if (selectedInput > 0 && selectedInput <= length(comb)) {
        for (i in 1:6) {
          required <- comb[[selectedInput]][[i]]
          inputed <- files[[i]]
          if (is.null(inputed) && is.null(required)) {
            reactiveVals$comparison[i] <- TRUE
          } else if (xor(is.null(inputed), is.null(required))) {
            reactiveVals$comparison[i] <- FALSE
          } else if (is.data.frame(required)) {
            reqCol <- colnames(inputed)
            inpCol <- required %>% pull(2)
            if (!identical(reqCol, inpCol)) {
              reactiveVals$comparison[i] <- FALSE
            } else {
              reactiveVals$comparison[i] <- TRUE
              colnames(files[[i]]) <- required %>% pull(1)
            }
          } else {
            reactiveVals$comparison[i] <- FALSE
          }
        }
      } else {
        reactiveVals$comparison <- c(FALSE, FALSE, FALSE, FALSE, FALSE, FALSE)
      }
    } else {
      reactiveVals$comparison <- c(FALSE, FALSE, FALSE, FALSE, FALSE, FALSE)
    }
    
    if (all(reactiveVals$comparison)) {
      enable("submit")
      removeClass("submit", "btn-disabled")
      addClass("submit", "btn-enabled")
    } else {
      disable("submit")
      removeClass("submit", "btn-enabled")
      addClass("submit", "btn-disabled")
    }
    
    output$curDmError <- renderUI({
      if (!is.null(files[[1]]) && reactiveVals$comparison[1] == FALSE) {
        errorSheet()
      } 
    })
    
    output$curEmpError <- renderUI({
      if (!is.null(files[[2]]) && reactiveVals$comparison[2] == FALSE) {
        errorSheet()
      } 
    })
    
    output$curStuError <- renderUI({
      if (!is.null(files[[3]]) && reactiveVals$comparison[3] == FALSE) {
        errorSheet()
      } 
    })
    
    output$preDmError <- renderUI({
      if (!is.null(files[[4]]) && reactiveVals$comparison[4] == FALSE) {
        errorSheet()
      } 
    })
    
    output$preEmpError <- renderUI({
      if (!is.null(files[[5]]) && reactiveVals$comparison[5] == FALSE) {
        errorSheet()
      } 
    })
    
    output$preStuError <- renderUI({
      if (!is.null(files[[6]]) && reactiveVals$comparison[6] == FALSE) {
        errorSheet()
      } 
    })
    
    reactiveVals$results <- files
  })
  
  observeEvent(input$submit, {
    if(!is.null(reactiveVals$results[[1]])){
      reactiveVals$results[[1]] <- reactiveVals$results[[1]] %>% mutate(Audience = "Decision-makers", Old = input$curOld, Survey = "Current") %>% select(Audience, Old, Survey, everything())
    }
    if(!is.null(reactiveVals$results[[2]])){
      reactiveVals$results[[2]] <- reactiveVals$results[[2]] %>% mutate(Audience = "Employees", Old = input$curOld, Survey = "Current") %>% select(Audience, Old, Survey, everything())
    }
    if(!is.null(reactiveVals$results[[3]])){
      reactiveVals$results[[3]] <- reactiveVals$results[[3]] %>% mutate(Audience = "Students", Old = input$curOld, Survey = "Current") %>% select(Audience, Old, Survey, everything())
    }
    if(!is.null(reactiveVals$results[[4]])){
      reactiveVals$results[[4]] <- reactiveVals$results[[4]] %>% mutate(Audience = "Decision-makers", Old = input$preOld, Survey = "Previous") %>% select(Audience, Old, Survey, everything())
    }
    if(!is.null(reactiveVals$results[[5]])){
      reactiveVals$results[[5]] <- reactiveVals$results[[5]] %>% mutate(Audience = "Employees", Old = input$preOld, Survey = "Previous") %>% select(Audience, Old, Survey, everything())
    }
    if(!is.null(reactiveVals$results[[6]])){
      reactiveVals$results[[6]] <- reactiveVals$results[[6]] %>% mutate(Audience = "Students", Old = input$preOld, Survey = "Previous") %>% select(Audience, Old, Survey, everything())
    }
    
    reactiveVals$dataWide <- bind_rows(compact(reactiveVals$results))
    reactiveVals$dataWide <- reactiveVals$dataWide[,order(colnames(reactiveVals$dataWide))]
    reactiveVals$dataWide <- reactiveVals$dataWide %>% select(Audience, Old, Survey, everything()) %>% arrange(Audience) %>% arrange(Survey) %>% mutate(ID = row_number(), .before = Audience)
    reactiveVals$dataLong <- reactiveVals$dataWide %>% pivot_longer(cols = -c(1:4), names_to = "Code", values_to = "Response") %>% filter(!is.na(Response))
    reactiveVals$dataLong <- reactiveVals$dataLong %>% left_join(master, by = "Code") %>% select(-colOld, -onFilter, -audFilter)
    reactiveVals$dataFactor <- left_join(reactiveVals$dataLong, factorLookup, by = "Response") %>% select(-rValue) %>% relocate(Order, .after = Response)
    reactiveVals$dataValue <- reactiveVals$dataLong %>% filter(!is.na(rValue)) %>% filter(!(rValue == 2 & Old == "Yes"))
    reactiveVals$dataValue <- left_join(reactiveVals$dataValue, valueLookup, by = "Response") %>% select(-rValue) %>% relocate(Value, .after = Response)
    
    if(input$previous == "Yes"){
      overviewData <- reactiveVals$dataFactor %>% filter(Survey == "Current")
      audData <- reactiveVals$dataWide %>% filter(Survey == "Current")
      reactiveVals$overviewSur <- reactiveVals$dataWide %>% group_by(Survey)
      reactiveVals$overviewSur <- overviewTable(reactiveVals$overviewSur) %>% select(1,5)
    } else if(input$previous == "No") {
      overviewData <- reactiveVals$dataFactor
      audData <- reactiveVals$dataWide
      reactiveVals$overviewSur <- NULL
    }
    reactiveVals$overviewAge <- overviewData %>% filter(gAna == "Age") %>% group_by(Response)
    reactiveVals$overviewGen <- overviewData %>% filter(gAna == "Gender") %>% group_by(Response)
    reactiveVals$overviewAud <- audData %>% group_by(Audience)
    reactiveVals$overviewAge <- overviewTable(reactiveVals$overviewAge) %>% select(1,5) %>% rename(Age = Response)
    reactiveVals$overviewGen <- overviewTable(reactiveVals$overviewGen) %>% select(1,5) %>% rename(Gender = Response)
    reactiveVals$overviewAud <- overviewTable(reactiveVals$overviewAud) %>% select(1,5)
    if(overviewData$Old[1] == "Yes"){
      overview147 <- overviewData %>% filter(Code == "Q147") %>% group_by(Order, Response)
      overview147 <- overviewTable(overview147) %>% select(2,6)
      overviewDataQs <- overviewData %>% filter(Code != "Q147") %>% filter(!is.na(q))
      nameQ <- NA
      colI <- "Item"
    } else if(overviewData$Old[1] == "No") {
      overview147 <- NULL
      overviewDataQs <- overviewData %>% filter(!is.na(q))
      nameQ <- "Please indicate how much you agree or disagree with the following statements"
      colI <- "Ind"
    }
    overviewDataDm <- overviewDataQs %>% filter(Audience == "Decision-makers")
    overviewDataUser <- overviewDataQs %>% filter(Audience != "Decision-makers")
    overviewDataAud <- list(`Decision-makers` = overviewDataDm, Users = overviewDataUser)
    reactiveVals$overviewQ <- list()
    for(aud in names(overviewDataAud)){
      overviewDataQ <- overviewDataAud[[aud]]
      reactiveVals$overviewQ[[aud]] <- list()
      for(qn in unique(overviewDataQ$q)){
        overviewQn <- overviewDataQ %>% filter(q == qn)
        nQn <- match(qn, unique(overviewDataQ$q))
        nameQn <- ifelse(is.na(nameQ), paste0(nQn, ". ", overviewQn$Question[1]), paste0(nQn, ". ", nameQ))
        overviewQn <- overviewQn %>% group_by(Code, across(all_of(colI)), Response, Order) %>%
          summarise(count = n(), .groups = 'drop') %>%
          group_by(Code, across(all_of(colI))) %>% mutate(N = sum(count)) %>% ungroup() %>%
          mutate(percent = count/N * 100, Count = paste(count, sprintf("(%.1f%%)", percent), sep = " ")) %>% arrange(Order)
        colnames(overviewQn)[2] <- "Question"
        overviewQnTable <- overviewQn %>% select(Code, Question, N, Response, Count)
        overviewQnTable <- pivot_wider(overviewQnTable, names_from = Response, values_from = Count) %>% arrange(Code) %>% select(-Code)
        overviewQnTable[is.na(overviewQnTable)] <- ""
        reactiveVals$overviewQ[[aud]][[nameQn]] <- overviewQnTable
      }
    }
    if(!is.null(overview147)){
      reactiveVals$overviewQ[[1]]$Q147 <- overview147
      reactiveVals$overviewQ[[1]] <- reactiveVals$overviewQ[[1]][c(1,2,3,4,5,7,6)]
      names(reactiveVals$overviewQ[[1]])[6] <- "6. How often does your organisation carry sustainability campaigns (e.g., turn it off)?"
      names(reactiveVals$overviewQ[[1]])[7] <- gsub("^6\\.", "7.", names(reactiveVals$overviewQ[[1]])[7])
    }
    
    if (input$previous == "Yes") {
      reactiveVals$choices <- c("Survey", "Audience", "Age", "Gender")
      reactiveVals$downloadDem <- list(reactiveVals$overviewSur, reactiveVals$overviewAud, reactiveVals$overviewAge, reactiveVals$overviewGen)
      names(reactiveVals$downloadDem) <- reactiveVals$choices
    } else if (input$previous == "No") {
      reactiveVals$choices <- c("Audience", "Age", "Gender")
      reactiveVals$downloadDem <- list(reactiveVals$overviewAud, reactiveVals$overviewAge, reactiveVals$overviewGen)
      names(reactiveVals$downloadDem) <- reactiveVals$choices
    }
    
    hData <- reactiveVals$dataValue %>% filter(!is.na(hAna) & Survey == "Current") %>% select(Code, hAna, Setting, Value)
    reactiveVals$hRes1 <- vector()
    reactiveVals$hRes2 <- vector()
    reactiveVals$hRes3 <- vector()
    for(question in unique(hData$hAna)){
      hPair <- hData %>% filter(hAna == question)
      hSig <- significance(wilcoxP(hPair, "Code"))
      hRes <- paste0(question, " (", hSig, ")")
      hTable <- hPair %>% group_by(Setting) %>% summarise(Score = round(mean(Value, na.rm = TRUE), digits = 2), .groups = 'drop') %>% arrange(desc(Score))
      if(hSig == "ns"){
        reactiveVals$hRes3 <- c(reactiveVals$hRes3, hRes)
      } else if(hTable[1,1] == "Home"){
        reactiveVals$hRes1 <- c(reactiveVals$hRes1, hRes)
      } else {
        reactiveVals$hRes2 <- c(reactiveVals$hRes2, hRes)
      }
    }
    reactiveVals$hRes1 <- blankVector(reactiveVals$hRes1)
    reactiveVals$hRes2 <- blankVector(reactiveVals$hRes2)
    reactiveVals$hRes3 <- blankVector(reactiveVals$hRes3)
    
    reactiveVals$comparisons <- list()
    maxLengthH <- max(length(reactiveVals$hRes1), length(reactiveVals$hRes2), length(reactiveVals$hRes3))
    vector1H <- c(reactiveVals$hRes1, rep("", maxLengthH - length(reactiveVals$hRes1)))
    vector2H <- c(reactiveVals$hRes2, rep("", maxLengthH - length(reactiveVals$hRes2)))
    vector3H <- c(reactiveVals$hRes3, rep("", maxLengthH - length(reactiveVals$hRes3)))
    dfH <- data.frame(vector1H, vector2H, vector3H)
    colnames(dfH) <- c("Home > Organisation", "Organisation > Home", "Home = Organisation")
    reactiveVals$comparisons[[1]] <- dfH
    
    reactiveVals$gRes1 <- vector()
    reactiveVals$gRes2 <- vector()
    reactiveVals$gRes3 <- vector()
    if(input$curAud == "Both"){
      if(input$curOld == "Yes"){
        gData <- reactiveVals$dataValue %>% filter(gAna == "1" & Survey == "Current") %>% select(Audience, Question, Item, Value) %>% mutate(Question = paste0(Question, ": ", Item)) %>% select(-Item)
      } else {
        gData <- reactiveVals$dataValue %>% filter(gAna == "1" | gAna == "2") %>% filter(Survey == "Current") %>% select(Audience, Ind, Value)
        colnames(gData)[2] <- "Question"
      }
      for(question in unique(gData$Question)){
        gPair <- gData %>% filter(Question == question)
        gSig <- significance(wilcoxP(gPair, "Audience"))
        gRes <- paste0(question, " (", gSig, ")")
        gTable <- gPair %>% group_by(Audience) %>% summarise(Score = round(mean(Value, na.rm = TRUE), digits = 2), .groups = 'drop') %>% arrange(desc(Score))
        if(gSig == "ns"){
          reactiveVals$gRes3 <- c(reactiveVals$gRes3, gRes)
        } else if(gTable[1,1] == "Students"){
          reactiveVals$gRes1 <- c(reactiveVals$gRes1, gRes)
        } else {
          reactiveVals$gRes2 <- c(reactiveVals$gRes2, gRes)
        }
      }
    }
    reactiveVals$gRes1 <- blankVector(reactiveVals$gRes1)
    reactiveVals$gRes2 <- blankVector(reactiveVals$gRes2)
    reactiveVals$gRes3 <- blankVector(reactiveVals$gRes3)
    
    maxLengthG <- max(length(reactiveVals$gRes1), length(reactiveVals$gRes2), length(reactiveVals$gRes3))
    vector1G <- c(reactiveVals$gRes1, rep("", maxLengthG - length(reactiveVals$gRes1)))
    vector2G <- c(reactiveVals$gRes2, rep("", maxLengthG - length(reactiveVals$gRes2)))
    vector3G <- c(reactiveVals$gRes3, rep("", maxLengthG - length(reactiveVals$gRes3)))
    dfG <- data.frame(vector1G, vector2G, vector3G)
    colnames(dfG) <- c("Students > Employees", "Employees > Students", "Students = Employees")
    reactiveVals$comparisons[[2]] <- dfG
    
    reactiveVals$tRes1 <- vector()
    reactiveVals$tRes2 <- vector()
    reactiveVals$tRes3 <- vector()
    if(input$previous == "Yes"){
      if(reactiveVals$preOld == "Yes"){
        tData <- reactiveVals$dataValue %>% filter(tAna == 1) %>% select(Survey, Ind, Value)
      } else {
        tData <- reactiveVals$dataValue %>% filter(!is.na(tAna)) %>% select(Survey, Ind, Value)
      }
      for(question in unique(tData$Ind)){
        tPair <- tData %>% filter(Ind == question)
        tSig <- significance(wilcoxP(tPair, "Survey"))
        tRes <- paste0(question, " (", tSig, ")")
        tTable <- tPair %>% group_by(Survey) %>% summarise(Score = round(mean(Value, na.rm = TRUE), digits = 2), .groups = 'drop') %>% arrange(desc(Score))
        if(tSig == "ns"){
          reactiveVals$tRes3 <- c(reactiveVals$tRes3, tRes)
        } else if(tTable[1,1] == "Current"){
          reactiveVals$tRes1 <- c(reactiveVals$tRes1, tRes)
        } else {
          reactiveVals$tRes2 <- c(reactiveVals$tRes2, tRes)
        }
      }
    }
    reactiveVals$tRes1 <- blankVector(reactiveVals$tRes1)
    reactiveVals$tRes2 <- blankVector(reactiveVals$tRes2)
    reactiveVals$tRes3 <- blankVector(reactiveVals$tRes3)
    
    maxLengthT <- max(length(reactiveVals$gRes1), length(reactiveVals$gRes2), length(reactiveVals$gRes3))
    vector1T <- c(reactiveVals$gRes1, rep("", maxLengthT - length(reactiveVals$gRes1)))
    vector2T <- c(reactiveVals$gRes2, rep("", maxLengthT - length(reactiveVals$gRes2)))
    vector3T <- c(reactiveVals$gRes3, rep("", maxLengthT - length(reactiveVals$gRes3)))
    dfT <- data.frame(vector1T, vector2T, vector3T)
    colnames(dfT) <- c("Current > Previous", "Previous > Current", "Current = Previous")
    reactiveVals$comparisons[[3]] <- dfT
    
    names(reactiveVals$comparisons) <- c("Setting", "Audience", "Time")
    
    if(input$previous == "No"){
      curModelData <- dataModel(reactiveVals$dataValue, "Current", input$curOld)
      preModelData <- NULL
    } else if(input$previous == "Yes") {
      curModelData <- dataModel(reactiveVals$dataValue, "Current", input$curOld)
      preModelData <- dataModel(reactiveVals$dataValue, "Previous", reactiveVals$preOld)
    }
    indResCur <- modelCalc(curModelData, "Current", "Ind", "Value") %>% rename(IndScore = Score, IndCat = Cat)
    critResCur <- modelCalc(indResCur, "Current", "Crit", "IndScore") %>% rename(CritScore = Score, CritCat = Cat)
    topicResCur <- modelCalc(critResCur, "Current", "Topic", "CritScore") %>% rename(TopicScore = Score, TopicCat = Cat)
    modelResCur <- topicResCur %>%
      left_join(critResCur, by = c("TopicCode", "Topic")) %>%
      left_join(indResCur, by = c("TopicCode", "Topic", "CritCode", "Crit"))
    if(input$previous == "Yes"){
      indResPre <- modelCalc(preModelData, "Previous", "Ind", "Value") %>% rename(IndScorePre = Score, IndCatPre = Cat)
      critResPre <- modelCalc(indResPre, "Previous", "Crit", "IndScorePre") %>% rename(CritScorePre = Score, CritCatPre = Cat)
      topicResPre <- modelCalc(critResPre, "Previous", "Topic", "CritScorePre") %>% rename(TopicScorePre = Score, TopicCatPre = Cat)
      modelRes <- modelResCur %>%
        left_join(topicResPre, by = c("TopicCode", "Topic")) %>% 
        left_join(critResPre, by = c("TopicCode", "Topic", "CritCode", "Crit")) %>% 
        left_join(indResPre, by = c("TopicCode", "Topic", "CritCode", "Crit", "IndCode", "Ind"))
      modelRes <- modelRes %>% percentChange("Topic") %>% percentChange("Crit") %>% percentChange("Ind")
      modelRes$TopicChangeLabel <- sapply(modelRes$TopicChange, changeLabel)
      modelRes$CritChangeLabel <- sapply(modelRes$CritChange, changeLabel)
      modelRes$IndChangeLabel <- sapply(modelRes$IndChange, changeLabel)
      modelRes <- modelRes %>% select(TopicCode, Topic, TopicScore, TopicCat, TopicScorePre, TopicCatPre, TopicChangeLabel,
                                       CritCode, Crit, CritScore, CritCat, CritScorePre, CritCatPre, CritChangeLabel,
                                       IndCode, Ind, IndScore, IndCat, IndScorePre, IndCatPre, IndChangeLabel)
      modelRes[] <- lapply(modelRes, function(x) {
        if (is.numeric(x)) {
          x <- as.character(x)
        }
        x[is.na(x)] <- ""
        return(x)
      })
      modelRes <- modelRes %>% select(-TopicCode, -CritCode, -IndCode)
      reactiveVals$indRes <- modelRes %>% select(Ind, IndScore, IndCat) %>% unique() %>%
        arrange(!is.na(IndScore), IndScore) %>% select(-IndScore)
      colnames(reactiveVals$indRes) <- c("Indicator", "Current Category")
      reactiveVals$critRes <- modelRes %>% select(Crit, CritScore, CritCat) %>% unique() %>%
        arrange(!is.na(CritScore), CritScore) %>% select(-CritScore)
      colnames(reactiveVals$critRes) <- c("Criterium", "Current Category")
      reactiveVals$topicRes <- modelRes %>% select(Topic, TopicScore, TopicCat) %>% unique() %>%
        arrange(!is.na(TopicScore), TopicScore) %>% select(-TopicScore)
      colnames(reactiveVals$topicRes) <- c("Topic", "Current Category")
    } else if(input$previous == "No"){
      modelRes <- modelResCur %>% select(-TopicCode, -CritCode, -IndCode)
      reactiveVals$indRes <- modelRes %>% select(Ind, IndScore, IndCat) %>% unique() %>%
        arrange(!is.na(IndScore), IndScore) %>% select(-IndScore)
      colnames(reactiveVals$indRes) <- c("Indicator", "Category")
      reactiveVals$critRes <- modelRes %>% select(Crit, CritScore, CritCat) %>% unique() %>%
        arrange(!is.na(CritScore), CritScore) %>% select(-CritScore)
      colnames(reactiveVals$critRes) <- c("Criterium", "Category")
      reactiveVals$topicRes <- modelRes %>% select(Topic, TopicScore, TopicCat) %>% unique() %>%
        arrange(!is.na(TopicScore), TopicScore) %>% select(-TopicScore)
      colnames(reactiveVals$topicRes) <- c("Topic", "Category")
    }
    
    if(input$previous == "Yes") {
      newDataPre <- list()
      levelsPre <- list()
      for (i in 1:nrow(modelRes)) {
        topicRowPre <- list(modelRes$Topic[i], modelRes$TopicScore[i], modelRes$TopicCat[i], modelRes$TopicChangeLabel[i], modelRes$TopicCatPre[i])
        criteriaRowPre <- list(modelRes$Crit[i], modelRes$CritScore[i], modelRes$CritCat[i], modelRes$CritChangeLabel[i], modelRes$CritCatPre[i])
        indicatorRowPre <- list(modelRes$Ind[i], modelRes$IndScore[i], modelRes$IndCat[i], modelRes$IndChangeLabel[i], modelRes$IndCatPre[i])
        if (length(newDataPre) == 0 || newDataPre[[length(newDataPre)]][1] != modelRes$Topic[i]) {
          newDataPre <- append(newDataPre, list(topicRowPre))
          levelsPre <- append(levelsPre, list(1))
        }
        newDataPre <- append(newDataPre, list(criteriaRowPre))
        levelsPre <- append(levelsPre, list(2))
        newDataPre <- append(newDataPre, list(indicatorRowPre))
        levelsPre <- append(levelsPre, list(3))
      }
      newDataDfPre <- do.call(rbind, lapply(newDataPre, function(x) data.frame(t(unlist(x)))))
      colnames(newDataDfPre) <- c("Description", "Score", "Category", "Score Change", "PreCat")
      newDataDfPre$Level <- unlist(levelsPre)
      newDataDfPre <- newDataDfPre %>% distinct()
      reactiveVals$modelResPrint <- newDataDfPre
      newDataDfPre$Description <- ifelse(
        newDataDfPre$Level == 2,
        cell_spec(newDataDfPre$Description, "html", extra_css = "margin-left: 20px; white-space: normal; display: block;"),
        ifelse(
          newDataDfPre$Level == 3,
          cell_spec(newDataDfPre$Description, "html", extra_css = "margin-left: 40px; white-space: normal; display: block;"),
          newDataDfPre$Description
        )
      )
      newDataDfDisplayPre <- newDataDfPre %>% select(-Level)
      reactiveVals$modelResPrint <- reactiveVals$modelResPrint %>% mutate(Type = case_when(Level == 1 ~ "Topic",
                                                                                           Level == 2 ~ "Criterium",
                                                                                           Level == 3 ~ "Indicator",
                                                                                           TRUE ~ NA
                                                                                           ), .before = Description
                                                                          ) %>% select(-Level)
      reactiveVals$modelResPrint$CatChange <- mapply(compareCat, reactiveVals$modelResPrint$Category, reactiveVals$modelResPrint$PreCat)
      colnames(reactiveVals$modelResPrint) <- c("Type", "Description", "Current Score", "Current Category", "Score Change", "PreCat", "Category Change")
      reactiveVals$modelResPrint <- reactiveVals$modelResPrint %>% select(-PreCat)
      newDataDfDisplayPre$CatChange <- mapply(compareCat, newDataDfDisplayPre$Category, newDataDfDisplayPre$PreCat)
      colnames(newDataDfDisplayPre) <- c("Description", "Current Score", "Current Category", "Score Change", "PreCat", "Category Change")
      newDataDfDisplayPre <- newDataDfDisplayPre %>% select(-PreCat)
      reactiveVals$modelRes <- newDataDfDisplayPre
      reactiveVals$modelResLevel <- newDataDfPre
      colnames(reactiveVals$modelRes) <- c("Description", "Current Score", "Current Category", "Score Change", "Category Change")
    } else if(input$previous == "No") {
      newData <- list()
      levels <- list()
      for (i in 1:nrow(modelRes)) {
        topicRow <- list(modelRes$Topic[i], modelRes$TopicScore[i], modelRes$TopicCat[i])
        criteriaRow <- list(modelRes$Crit[i], modelRes$CritScore[i], modelRes$CritCat[i])
        indicatorRow <- list(modelRes$Ind[i], modelRes$IndScore[i], modelRes$IndCat[i])
        if (length(newData) == 0 || newData[[length(newData)]][1] != modelRes$Topic[i]) {
          newData <- append(newData, list(topicRow))
          levels <- append(levels, list(1))
        }
        newData <- append(newData, list(criteriaRow))
        levels <- append(levels, list(2))
        newData <- append(newData, list(indicatorRow))
        levels <- append(levels, list(3))
      }
      newDataDf <- do.call(rbind, lapply(newData, function(x) data.frame(t(unlist(x)))))
      colnames(newDataDf) <- c("Description", "Score", "Category")
      newDataDf$Level <- unlist(levels)
      newDataDf <- newDataDf %>% distinct()
      reactiveVals$modelResPrint <- newDataDf
      newDataDf$Description <- ifelse(
        newDataDf$Level == 2,
        cell_spec(newDataDf$Description, "html", extra_css = "margin-left: 20px; white-space: normal; display: block;"),
        ifelse(
          newDataDf$Level == 3,
          cell_spec(newDataDf$Description, "html", extra_css = "margin-left: 40px; white-space: normal; display: block;"),
          newDataDf$Description
        )
      )
      newDataDfDisplay <- newDataDf %>% select(-Level)
      
      reactiveVals$modelResPrint <- reactiveVals$modelResPrint %>% mutate(Type = case_when(Level == 1 ~ "Topic",
                                                                                           Level == 2 ~ "Criterium",
                                                                                           Level == 3 ~ "Indicator",
                                                                                           TRUE ~ NA
                                                                                           ), .before = Description
                                                                          ) %>% select(-Level)
      colnames(reactiveVals$modelResPrint) <- c("Type", "Description", "Score", "Category")
      reactiveVals$modelRes <- newDataDfDisplay
      reactiveVals$modelResLevel <- newDataDf
      colnames(reactiveVals$modelRes) <- c("Description", "Score", "Category")
    }
  })
  
  observeEvent(input$showMenu, {
    if (!rv$showMenu) {
      rv$showMenu <- TRUE
      updateTabItems(session, "sidebarMenu", "dataUpload")
    }
  })
  
  observeEvent(input$submit, {
    if (!rv$showResult) {
      rv$showResult <- TRUE
      updateTabItems(session, "sidebarMenu", "results")
    }
  })
  
  output$dynamicMenu1 <- renderUI({
    if (rv$showMenu) {
      sidebarMenu(
        id = "sidebarMenu",
        menuItem("Data Upload", tabName = "dataUpload", icon = icon("upload"))
      )
    }
  })
  
  output$dynamicMenu2 <- renderUI({
    if (rv$showResult) {
      sidebarMenu(
        id = "sidebarMenu",
        menuItem("Results", tabName = "results", icon = icon("chart-line")),
        menuItem("Highlights", tabName = "highlights", icon = icon("star"))
      )
    }
  })
  
  output$selectOverviewDem <- renderUI({
    if(input$previous == "Yes"){
      selectInput("selectInputOverviewDem", label = "Select demographic",
                  choices = c("Survey", "Audience", "Age", "Gender"),
                  selected = "Survey"
      )
    } else if(input$previous == "No"){
      selectInput("selectInputOverviewDem", label = "Select demographic",
                  choices = c("Audience", "Age", "Gender"),
                  selected = "Audience"
      )
    }
  })
  
  output$selectOverviewDm <- renderUI({
    selectInput("selectInputOverviewDm", label = "Select question",
                choices = c(1:length(reactiveVals$overviewQ[[1]])),
                selected = 1
    )
  })
  
  output$selectOverviewUser <- renderUI({
    selectInput("selectInputOverviewUser", label = "Select question",
                choices = c(1:length(reactiveVals$overviewQ[[2]])),
                selected = 1
    )
  })
  
  output$overviewDem <- renderUI({
    req(input$selectInputOverviewDem)
    req(reactiveVals$choices)
    req(reactiveVals$downloadDem)
    if (!is.null(input$selectInputOverviewDem)) {
      selectedIndex <- which(reactiveVals$choices == input$selectInputOverviewDem)
      if (length(selectedIndex) > 0 && length(reactiveVals$downloadDem) >= selectedIndex) {
        selectedElement <- reactiveVals$downloadDem[[selectedIndex]]
        if (!is.null(selectedElement) && is.data.frame(selectedElement)) {
          HTML(
            selectedElement %>%
              kable() %>%
              kable_styling() %>%
              row_spec(0, extra_css = "border-bottom: 2px solid black;") %>%
              row_spec(0:nrow(selectedElement), extra_css = "padding: 2px 4px;")
          )
        }
      }
    }
  })
  
  output$overviewDm <- renderUI({
    req(input$selectInputOverviewDm)
    req(reactiveVals$overviewQ)
    x <- as.numeric(input$selectInputOverviewDm)
    HTML(
      reactiveVals$overviewQ[[1]][[x]] %>%
        kable(caption = names(reactiveVals$overviewQ[[1]])[x]) %>%
        kable_styling() %>%
        row_spec(0, extra_css = "border-bottom: 2px solid black;") %>%
        row_spec(0:nrow(reactiveVals$overviewQ[[1]][[x]]), extra_css = "padding: 2px 4px;")
    )
  })
  
  output$overviewUser <- renderUI({
    req(input$selectInputOverviewUser)
    req(reactiveVals$overviewQ)
    x <- as.numeric(input$selectInputOverviewUser)
    HTML(
      reactiveVals$overviewQ[[2]][[x]] %>%
        kable(caption = names(reactiveVals$overviewQ[[2]])[x]) %>%
        kable_styling() %>%
        row_spec(0, extra_css = "border-bottom: 2px solid black;") %>%
        row_spec(0:nrow(reactiveVals$overviewQ[[2]][[x]]), extra_css = "padding: 2px 4px;")
    )
  })
  
  output$hAna <- renderUI({
    req(reactiveVals$comparisons)
    HTML(
      reactiveVals$comparisons[[1]] %>%
        kable(caption = "Difference in score (home vs. organisation)", col.names = c("Home > Organisation", "Organisation > Home", "Home = Organisation")) %>%
        kable_styling() %>%
        row_spec(0, extra_css = "border-bottom: 2px solid black;") %>%
        row_spec(0:nrow(reactiveVals$comparisons[[1]]), extra_css = "padding: 2px 4px;") %>%
        collapse_rows(columns = 1:3, valign = "middle")
    )
  })
  
  output$gAna <- renderUI({
    req(reactiveVals$comparisons)
    if(input$curAud == "Both") {
      HTML(
        reactiveVals$comparisons[[2]] %>%
          kable(caption = "Difference in score (students vs. employees)", col.names = c("Students > Employees", "Employees > Students", "Students = Employees")) %>%
          kable_styling() %>%
          row_spec(0, extra_css = "border-bottom: 2px solid black;") %>%
          row_spec(0:nrow(reactiveVals$comparisons[[2]]), extra_css = "padding: 2px 4px;") %>%
          collapse_rows(columns = 1:3, valign = "middle")
      )
    } else {
      "No results to display. There is a single audience, so it is not possible to compare audiences."
    }
  })
  
  output$tAna <- renderUI({
    req(reactiveVals$comparisons)
    if(input$previous == "Yes") {
      HTML(
        reactiveVals$comparisons[[3]] %>%
          kable(caption = "Difference in score (current vs. previous)", col.names = c("Current > Previous", "Previous > Current", "Current = Previous")) %>%
          kable_styling() %>%
          row_spec(0, extra_css = "border-bottom: 2px solid black;") %>%
          row_spec(0:nrow(reactiveVals$comparisons[[3]]), extra_css = "padding: 2px 4px;") %>%
          collapse_rows(columns = 1:3, valign = "middle")
      )
    } else {
      "No results to display. There are no previous results, so it is not possible to track changes over time."
    }
  })
  
  output$modelResults <- renderUI({
    req(reactiveVals$modelRes)
    req(reactiveVals$modelResLevel)
    HTML(
      kable(reactiveVals$modelRes, "html", escape = FALSE, caption = "Model results (topics in dark grey, criteria in light grey, and indicators in white)") %>%
        kable_styling() %>%
        row_spec(which(reactiveVals$modelResLevel$Level == 1), bold = TRUE, background = "#bbbbbb") %>%
        row_spec(which(reactiveVals$modelResLevel$Level == 2), bold = TRUE, background = "#dddddd") %>%
        row_spec(which(reactiveVals$modelResLevel$Level == 3), background = "white") %>%
        row_spec(0, extra_css = "border-top: 1px solid black; border-left: 1px solid black; border-right: 1px solid black; border-bottom: 2px solid black;") %>%
        column_spec(1:ncol(reactiveVals$modelRes), extra_css = "border: 1px solid black;") %>%
        row_spec(0:nrow(reactiveVals$modelResLevel), extra_css = "padding: 2px 4px;")
    )
  })
  
  output$critRes <- renderUI({
    req(reactiveVals$critRes)
    HTML(
      reactiveVals$critRes %>%
        kable(caption = "Criteria arranged from lower scores (top) to higher scores (bottom). Criteria with no scores (NA) are showed on top. 
              Items on the top of the list are to be targeted first.") %>%
        kable_styling() %>%
        row_spec(which(reactiveVals$critRes[,2] == "A"), background = "#66CA64") %>%
        row_spec(which(reactiveVals$critRes[,2] == "B"), background = "#99D970") %>%
        row_spec(which(reactiveVals$critRes[,2] == "C"), background = "#FFD015") %>%
        row_spec(which(reactiveVals$critRes[,2] == "D"), background = "#FE8643") %>%
        row_spec(which(reactiveVals$critRes[,2] == "E"), background = "#FE5660") %>%
        row_spec(0, extra_css = "border-bottom: 2px solid black;") %>%
        row_spec(0:nrow(reactiveVals$critRes), extra_css = "padding: 2px 4px;")
    )
  })
  
  output$indRes <- renderUI({
    req(reactiveVals$indRes)
    HTML(
      reactiveVals$indRes %>%
        kable(caption = "Indicators arranged from lower scores (top) to higher scores (bottom). Indicators with no scores (NA) are showed on top. 
              Items on the top of the list are to be targeted first.") %>%
        kable_styling() %>%
        row_spec(which(reactiveVals$indRes[,2] == "A"), background = "#66CA64") %>%
        row_spec(which(reactiveVals$indRes[,2] == "B"), background = "#99D970") %>%
        row_spec(which(reactiveVals$indRes[,2] == "C"), background = "#FFD015") %>%
        row_spec(which(reactiveVals$indRes[,2] == "D"), background = "#FE8643") %>%
        row_spec(which(reactiveVals$indRes[,2] == "E"), background = "#FE5660") %>%
        row_spec(0, extra_css = "border-bottom: 2px solid black;") %>%
        row_spec(0:nrow(reactiveVals$indRes), extra_css = "padding: 2px 4px;")
    )
  })
  
  output$topicRes <- renderUI({
    req(reactiveVals$topicRes)
    HTML(
      reactiveVals$topicRes %>%
        kable(caption = "Topics arranged from lower scores (top) to higher scores (bottom). Topics with no scores (NA) are showed on top. 
              Items on the top of the list are to be targeted first.") %>%
        kable_styling() %>%
        row_spec(which(reactiveVals$topicRes[,2] == "A"), background = "#66CA64") %>%
        row_spec(which(reactiveVals$topicRes[,2] == "B"), background = "#99D970") %>%
        row_spec(which(reactiveVals$topicRes[,2] == "C"), background = "#FFD015") %>%
        row_spec(which(reactiveVals$topicRes[,2] == "D"), background = "#FE8643") %>%
        row_spec(which(reactiveVals$topicRes[,2] == "E"), background = "#FE5660") %>%
        row_spec(0, extra_css = "border-bottom: 2px solid black;") %>%
        row_spec(0:nrow(reactiveVals$topicRes), extra_css = "padding: 2px 4px;")
    )
  })
  
  output$modelDescription <- renderUI({
    HTML(
      modelMeaning %>%
        kable(caption = "UI-EM3 structure") %>%
        kable_styling() %>%
        row_spec(0, extra_css = "border-bottom: 2px solid black;") %>%
        row_spec(0:nrow(modelMeaning), extra_css = "padding: 2px 4px;") %>%
        collapse_rows(columns = 1:2, valign = "middle")
    )
  })
  
  output$catDescription <- renderUI({
    HTML(
      catMeaning %>%
        mutate(
          Category = case_when(
            row_number() == 1 ~ cell_spec(Category, background = "#66CA64"),
            row_number() == 2 ~ cell_spec(Category, background = "#99D970"),
            row_number() == 3 ~ cell_spec(Category, background = "#FFD015"),
            row_number() == 4 ~ cell_spec(Category, background = "#FE8643"),
            row_number() == 5 ~ cell_spec(Category, background = "#FE5660"),
            TRUE ~ Category
          )
        ) %>%
        kable(format = "html", caption = "Maturity categories", escape = FALSE) %>%
        kable_styling() %>%
        row_spec(0, extra_css = "border-bottom: 2px solid black;") %>%
        row_spec(0:nrow(catMeaning), extra_css = "padding: 2px 4px;") %>%
        collapse_rows(columns = 1:2, valign = "middle")
    )
  })
  
  output$sigDescription <- renderUI({
    HTML(
      sigMeaning %>%
        kable(caption = "Significance description") %>%
        kable_styling() %>%
        row_spec(0, extra_css = "border-bottom: 2px solid black;") %>%
        row_spec(0:nrow(sigMeaning), extra_css = "padding: 2px 4px;") %>%
        collapse_rows(columns = 1:2, valign = "middle")
    )
  })
  
  output$downloadDm <- downloadHandler(
    filename = function() {
      "Example Decision-makers.xlsx"
    },
    content = function(file) {
      file.copy("Example Decision-makers.xlsx", file)
    }
  )
  
  output$downloadEmp <- downloadHandler(
    filename = function() {
      "Example Employees.xlsx"
    },
    content = function(file) {
      file.copy("Example Employees.xlsx", file)
    }
  )
  
  output$downloadStuUni <- downloadHandler(
    filename = function() {
      "Example University Students.xlsx"
    },
    content = function(file) {
      file.copy("Example University Students.xlsx", file)
    }
  )
  
  output$downloadStuSch <- downloadHandler(
    filename = function() {
      "Example School Students.xlsx"
    },
    content = function(file) {
      file.copy("Example School Students.xlsx", file)
    }
  )
  
  output$downloadDem <- downloadHandler(
    filename = function() { "Demographics Overview.xlsx" },
    content = function(file) {
      writeListExcel(reactiveVals$downloadDem, file)
    }
  )
  
  output$downloadDM <- downloadHandler(
    filename = function() { "Decision-makers Overview.xlsx" },
    content = function(file) {
      writeListCaption(reactiveVals$overviewQ[[1]], file)
    }
  )
  
  output$downloadUser <- downloadHandler(
    filename = function() { "Users Overview.xlsx" },
    content = function(file) {
      writeListCaption(reactiveVals$overviewQ[[2]], file)
    }
  )
  
  output$downloadComparison <- downloadHandler(
    filename = function() { "Comparisons.xlsx" },
    content = function(file) {
      writeListExcel(reactiveVals$comparisons, file)
    }
  )
  
  output$downloadModel <- downloadHandler(
    filename = function() { "Model Results.xlsx" },
    content = function(file) {
      writeDataframeExcel(file)
    }
  )
}

# Run the app ----
shinyApp(ui, server)
