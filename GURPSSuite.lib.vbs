'author Alan Bartholet

'Includes
Include "GURPSSuite.conf.vbs"

'------------------------------------------------------------------------------
'Build the System
'------------------------------------------------------------------------------
Sub BuildSystem(body, buildChildren)
	'Declare variables
	Dim stellarAge    'as Double

	'If this is the root body or the root body doesn't doesn't have a stellar age
	'we need to set the stellar age
	If Not body.HasParent() Or GetGURPSValue(body.GetRootBody, "stellar_age") = "" Then
		SetStellarAge body.GetRootBody, GetStellarAge
	End If

	'Build Terrestrial Body
	If body.TypeID = BODY_TYPE_TERRESTRIAL Then
		BuildTerrestrialBody body

	'Build Gas Giant Body
	ElseIf body.TypeID = BODY_TYPE_GASGIANT Then
		'Reset the sulfur flag
		If buildChildren = TRUE Then
			SetGURPSValue body, "gas_sulfur", ""
		End If
		BuildGasGiantBody body

	'Build Asteroid Belt Body
	ElseIf body.TypeID = BODY_TYPE_ASTEROIDBELT Then
		BuildAsteroidBeltBody body

	'Build Small Body
	ElseIf body.TypeID = BODY_TYPE_SMALLBODY Then
		BuildSmallBody body

	'Build Planetoid Body
	ElseIf body.TypeID = BODY_TYPE_PLANETOID Then
		BuildPlanetoidBody body
	End If

	'Loop through each child body
	If body.ChildrenCount() > 0 And buildChildren = TRUE Then
		For i = 1 To body.ChildrenCount()
			'Build each child body
			BuildSystem body.GetChild(i - 1), buildChildren
		Next
	End If
End Sub


'------------------------------------------------------------------------------
'Build Star Body
'------------------------------------------------------------------------------
Sub BuildStarBody(body)

	'Display the body
	DisplayStarBody(body)
End Sub


'------------------------------------------------------------------------------
'Build Asteroid Belt Body
'------------------------------------------------------------------------------
Sub BuildAsteroidBeltBody(body)
	'Build the resource value
	BuildResourceValue(body)

	'Build the habitability value
	BuildHabitabilityValue(body)

	'Display the body
	DisplayAsteroidBeltBody(body)
End Sub


'------------------------------------------------------------------------------
'Build Planetoid Body
'------------------------------------------------------------------------------
Sub BuildPlanetoidBody(body)
	'Build the climate type
	BuildClimateType(body)

	'Build the resource value
	BuildResourceValue(body)

	'Build the habitability value
	BuildHabitabilityValue(body)

	'Display the body
	DisplayPlanetoidBody(body)
End Sub


'------------------------------------------------------------------------------
'Build Small Body
'------------------------------------------------------------------------------
Sub BuildSmallBody(body)
	If body.Distance = 0 Then
		'According to NBOS Support
		''Small bodies' is really just a note that there are numerous small,
		'unremarkable bodies in orbit (small chunks of rock).  Its not a single
		'body with a single orbit, but rather debris found in orbit at various
		'places.  So there's no orbital parameters per se, and they arent shown in
		'the system diagram.  Think of it as the orbital equivalent of a 'look out
		'for falling rocks' sign.
		'Unfortunately by default 'Small Bodies' have a distance of 0. This causes the Travel Calculator
		'to throw an invalid floating point violation error, in order to fix this we will set
		'the distance to 0.01.  This will achieve the same affect for all intents and purposes
		'without causing the error.
		body.Distance = 0.01
		unloadBody = FALSE
	End If

	'Set the version that this body was created with
	SetGURPSValue body, "suite_version", GURPSSuiteVersion
End Sub


'------------------------------------------------------------------------------
'Build Gas Giant Body
'------------------------------------------------------------------------------
Sub BuildGasGiantBody(body)
	'Determine the size classification of the gas giant
	If body.Mass < 100 Then
		SetGURPSValue body, "gas_size", 0
	ElseIf body.Mass < 600 Then
		SetGURPSValue body, "gas_size", 1
	Else
		SetGURPSValue body, "gas_size", 2
	End If

	'Display the body
	DisplayGasGiantBody(body)
End Sub


'------------------------------------------------------------------------------
'Build Terrestrial Planet
'------------------------------------------------------------------------------
Sub BuildTerrestrialBody(body)
	'Declare variables
	Dim tmp, tmp2                       'as Double
	Dim blackBody, mmwr                 'as Int
	Dim sizeMult, eRadius, Ts, Rs, D    'as Double

	'Calculate the black body temperature based on the parent star
	If Not IsNull(body.GetParentStar) Then
		Ts = (1/(((body.GetParentStar.Radius/1)/(body.GetParentStar.Luminosity)^(1/2))^(1/2)) * 5780)
		Rs = body.GetParentStar.Radius * (6.96*(10^8))
		D = body.DistanceFromParentStar * 1000
		StarMass = body.GetParentStar.Mass
	Else
		'If there is no parent star we will use the planets surface temperature.  If the surface
		'temperature is less than 2.728 Kelvin, the generally accepted "temperature" of deep space,
		'we will use that instead
		Rs = 0.1
		D = 0.1
		If body.Temp < 2.728 Then
			Ts = 2.728
		Else
			body.Temp
		End If
		StarMass = 0
	End If
	blackBody = Ts * (((((1-body.Albedo)^(1/2)) * Rs)/D*2)^(1/2))/2

	'Get the plant's radius in Earths
	eRadius = body.Radius/6378

	'Calculate the planets minimum molecular weight retained
	mmwr = blackBody/(60 * (eRadius)^2 * body.Density)

	'Get the planet size based on the it's minimum molecular weight retained
	'*************************'
	' Tiny                    '
	'*************************'
	If mmwr > 28 Then
		SetGURPSValue body, "terrestrial_size", 0    'Tiny
		If body.Temp <= 140 Then
			If Not IsNull(body.GetParentBody) Then
				'Tiny sulfur planets appear as moons of gas giants
				'There is a 50/50 chance that a gas giant will have a sulfur planet
				'If the gas giant already has a sulfur moon then all the others should be ice
				'If the sulfur flag is an empty string then this gas giant has not been checked
				'for a sulfur planet yet otherwise it has been checked
				'This shoud happen for the first inner most moon that is applicable
				If body.GetParentBody.TypeID = BODY_TYPE_GASGIANT And GetGURPSValue(body.GetParentBody, "gas_sulfur") = "" And RollDice(1,6,0) <= 3 Then
					SetGURPSValue body, "terrestrial_type", 0           'Sulfur
					SetGURPSValue body, "atmospheric_composition", 0    'No/Negligible Composition
					SetGURPSValue body.GetParentBody, "gas_sulfur", 1
				Else
					SetGURPSValue body, "terrestrial_type", 1           'Ice
					SetGURPSValue body, "atmospheric_composition", 0    'No/Negligible Composition
					SetGURPSValue body.GetParentBody, "gas_sulfur", 0
				End If
			Else
				SetGURPSValue body, "terrestrial_type", 1               'Ice
				SetGURPSValue body, "atmospheric_composition", 0        'No/Negligible Composition
			End If
		Else
			SetGURPSValue body, "terrestrial_type", 2                   'Rock
			SetGURPSValue body, "atmospheric_composition", 0            'No/Negligible Composition
		End If

	'*************************'
	' Small                   '
	'*************************'
	ElseIf mmwr > 18 Then
		SetGURPSValue body, "terrestrial_size", 1    'Small

		If body.Temp <= 80 Then
			SetGURPSValue body, "terrestrial_type", 3                'Hadean
			SetGURPSValue body, "atmospheric_composition", 0         'No/Negligible Composition
		ElseIf body.Temp <= 140 Then
			SetGURPSValue body, "terrestrial_type", 1                'Ice
			If RollDice(3,6,0) <= 15 Then
				SetGURPSValue body, "atmospheric_composition", 13    'Suffocating and Mildly Toxic
			Else
				SetGURPSValue body, "atmospheric_composition", 14    'Suffocating and Highly Toxic
			End If
		Else
			SetGURPSValue body, "terrestrial_type", 2                'Rock
			SetGURPSValue body, "atmospheric_composition", 0         'No/Negligible Composition
		End If

	'*************************'
	' Standard                '
	'*************************'
	ElseIf mmwr > 4 Then
		SetGURPSValue body, "terrestrial_size", 2    'Standard

		If body.Temp <= 80 Then
			SetGURPSValue body, "terrestrial_type", 3                'Hadean
			SetGURPSValue body, "atmospheric_composition", 0         'No/Negligible Composition
		ElseIf body.Temp <= 230 And StarMass <= 0.65 Then
			SetGURPSValue body, "terrestrial_type", 4                'Ammonia
			SetGURPSValue body, "atmospheric_composition", 15        'Suffocating, Lethally Toxic, and Corrosive

		ElseIf body.Temp <= 240 Then
			SetGURPSValue body, "terrestrial_type", 1                'Ice
			If RollDice(3,6,0) <= 15 Then
				SetGURPSValue body, "atmospheric_composition", 12    'Suffocating
			Else
				SetGURPSValue body, "atmospheric_composition", 13    'Suffocating and Mildly Toxic
			End If

		ElseIf body.Temp <= 320 Then
			'This can be an Ocean or Garden planet depending on the age
			'of the world and the star system
			'First we pick a random number between 3 and 18
			tmp = RollDice(3,6,0)

			'Next we get a +1 for every full 500 million years old the
			'star is to a maximum of +10 for standard or +5 for large
			tmp2 = Int((GetGURPSValue(body.GetRootBody, "stellar_age") * 10^9)/(5 * 10^8))

			'If we are over the maxBonus set it to be the maxBonus
			If tmp2 > 10 Then
				tmp2 = 10
			End If

			'Add the two values together and if it is 18 or greater
			'the planet will be a Garden otherwise it is an ocean
			If (tmp + tmp2) < 18 Then
				SetGURPSValue body, "terrestrial_type", 5                'Ocean
				If RollDice(3,6,0) <= 15 Then
					SetGURPSValue body, "atmospheric_composition", 12    'Suffocating
				Else
					SetGURPSValue body, "atmospheric_composition", 13    'Suffocating and Mildly Toxic
				End If

			Else
				SetGURPSValue body, "terrestrial_type", 6               'Garden
				If RollDice(3,6,0) <= 11 Then
					SetGURPSValue body, "atmospheric_composition", 1    'Breathable
				Else
					BuildMarginalAtmosphere body                        'Marginal
				End If

			End If
		Else
			If body.Atmosphere > 0 Then
				SetGURPSValue body, "terrestrial_type", 7            'Greenhouse
				SetGURPSValue body, "atmospheric_composition", 15    'Suffocating, Lethally Toxic, and Corrosive

			Else
				SetGURPSValue body, "terrestrial_type", 8            'Chthonian
				SetGURPSValue body, "atmospheric_composition", 0     'No/Negligible Composition

			End If
		End If

	'*************************'
	' Large                   '
	'*************************'
	Else
		SetGURPSValue body, "terrestrial_size", 3    'Large

		If body.Temp <= 230 And StarMass <= 0.65 Then
			SetGURPSValue body, "terrestrial_type", 4            'Ammonia
			SetGURPSValue body, "atmospheric_composition", 15    'Suffocating, Lethally Toxic, and Corrosive
		ElseIf body.Temp <= 240 Then
			SetGURPSValue body, "terrestrial_type", 1            'Ice
			SetGURPSValue body, "atmospheric_composition", 14    'Suffocating and Highly Toxic

		ElseIf body.Temp <= 320 Then
			'This can be an Ocean or Garden planet depending on the age
			'of the world and the star system
			'First we pick a random number between 3 and 18
			tmp = RollDice(3,6,0)

			'Next we get a +1 for every full 500 million years old the
			'star is to a maximum of +10 for standard or +5 for large
			tmp2 = Int((GetGURPSValue(body.GetRootBody, "stellar_age") * 10^9)/(5 * 10^8))

			'If we are over the maxBonus set it to be the maxBonus
			If tmp2 > 5 Then
				tmp2 = 5
			End If

			'Add the two values together and if it is 18 or greater
			'the planet will be a Garden otherwise it is an ocean
			If (tmp + tmp2) < 18 Then
				SetGURPSValue body, "terrestrial_type", 5            'Ocean
				SetGURPSValue body, "atmospheric_composition", 13    'Suffocating and Highly Toxic

			Else
				SetGURPSValue body, "terrestrial_type", 6               'Garden
				If RollDice(3,6,0) <= 11 Then
					SetGURPSValue body, "atmospheric_composition", 1    'Breathable
				Else
					BuildMarginalAtmosphere body                        'Marginal
				End If

			End If
		Else
			If body.Atmosphere > 0 Then
				SetGURPSValue body, "terrestrial_type", 7            'Greenhouse
				SetGURPSValue body, "atmospheric_composition", 15    'Suffocating, Lethally Toxic, and Corrosive
			Else
				SetGURPSValue body, "terrestrial_type", 8            'Chthonian
				SetGURPSValue body, "atmospheric_composition", 0     'No/Negligible Composition
			End If
		End If
	End If



	'Build the climate type
	BuildClimateType(body)

	'Build the volcanic activity
	BuildVolcanicActivity(body)

	'Build the tectonic activity
	BuildTectonicActivity(body)

	'Build Atmosphere and Hydrographic Coverage
	BuildAtmosphereAndHydrographicCoverage(body)

	'Build the atmospheric density
	BuildAtmosphereDensity(body)

	'Build the resource value
	BuildResourceValue(body)

	'Build the habitability score
	BuildHabitabilityValue(body)

	'Display the body
	DisplayTerrestrialBody(body)
End Sub


'------------------------------------------------------------------------------
'Build Atmosphere and Hydrographic Coverage
'------------------------------------------------------------------------------
Sub BuildAtmosphereAndHydrographicCoverage(body)
	Select Case GetGURPSValue(body, "terrestrial_size")
		''''''''
		' Tiny '
		''''''''
		Case 0
			Select Case GetGURPSValue(body, "terrestrial_type")
				''''''''''
				' Sulfur '
				''''''''''
				Case 0
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = 0
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = 0
					End If

				'''''''
				' Ice '
				'''''''
				Case 1
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = 0
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = 0
					End If

				''''''''
				' Rock '
				''''''''
				Case 2
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = 0
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = 0
					End If

			End Select

		'''''''''
		' Small '
		'''''''''
		Case 1
			Select Case GetGURPSValue(body, "terrestrial_type")
				'''''''
				' Ice '
				'''''''
				Case 1
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = GetHydrographicCoverage(1, 2)
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = GetAtmosphericPressure(body, 10)
					End If
					Select Case GetGURPSValue(body, "atmospheric_composition")
						''''''''''''''''''''''''''''''''
						' Suffocating and Mildly Toxic '
						''''''''''''''''''''''''''''''''
						Case 13
							body.SetAtmElement "CH4", GetAtmosphericGasQuantity(body, 1, 5)
							body.SetAtmElement "N2", GetAtmosphericGasQuantity(body, 100, 100)

						''''''''''''''''''''''''''''''''
						' Suffocating and Highly Toxic '
						''''''''''''''''''''''''''''''''
						Case 14
							body.SetAtmElement "CH4", GetAtmosphericGasQuantity(body, 1, 5)
							body.SetAtmElement "N2", GetAtmosphericGasQuantity(body, 100, 100)
					End Select

				''''''''
				' Rock '
				''''''''
				Case 2
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = 0
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = Rand(1, 10) / 1000
					End If

				''''''''''
				' Hadean '
				''''''''''
				Case 3
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = 0
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = 0
					End If

			End Select

		''''''''''''
		' Standard '
		''''''''''''
		Case 2
			Select Case GetGURPSValue(body, "terrestrial_type")
				'''''''
				' Ice '
				'''''''
				Case 1
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = GetHydrographicCoverage(2, -10)
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = GetAtmosphericPressure(body, 1)
					End If
					Select Case GetGURPSValue(body, "atmospheric_composition")
						'''''''''''''''
						' Suffocating '
						'''''''''''''''
						Case 12
							body.SetAtmElement "C02", GetAtmosphericGasQuantity(body, 20, 40)
							body.SetAtmElement "N2", GetAtmosphericGasQuantity(body, 100, 100)

						''''''''''''''''''''''''''''''''
						' Suffocating and Mildly Toxic '
						''''''''''''''''''''''''''''''''
						Case 13
							body.SetAtmElement "C02", GetAtmosphericGasQuantity(body, 20, 40)
							body.SetAtmElement "N2", GetAtmosphericGasQuantity(body, 100, 100)
					End Select

				''''''''''
				' Hadean '
				''''''''''
				Case 3
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = 0
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = 0
					End If

				'''''''''''
				' Ammonia '
				'''''''''''
				Case 4
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = GetHydrographicCoverage(2, 0)
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = GetAtmosphericPressure(body, 1)
					End If
					Select Case GetGURPSValue(body, "atmospheric_composition")
						''''''''''''''''''''''''''''''''''''''''''''''
						' Suffocating, Lethally Toxic, and Corrosive '
						''''''''''''''''''''''''''''''''''''''''''''''
						Case 15
							body.SetAtmElement "NH3", GetAtmosphericGasQuantity(body, 1, 20)
							body.SetAtmElement "CH4", GetAtmosphericGasQuantity(body, 1, 20)
							body.SetAtmElement "N2", GetAtmosphericGasQuantity(body, 100, 100)
					End Select

				'''''''''
				' Ocean '
				'''''''''
				Case 5
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = GetHydrographicCoverage(1, 4)
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = GetAtmosphericPressure(body, 1)
					End If
					Select Case GetGURPSValue(body, "atmospheric_composition")
						'''''''''''''''
						' Suffocating '
						'''''''''''''''
						Case 12
							body.SetAtmElement "C02", GetAtmosphericGasQuantity(body, 20, 40)
							body.SetAtmElement "N2", GetAtmosphericGasQuantity(body, 100, 100)

						''''''''''''''''''''''''''''''''
						' Suffocating and Mildly Toxic '
						''''''''''''''''''''''''''''''''
						Case 13
							body.SetAtmElement "C02", GetAtmosphericGasQuantity(body, 20, 40)
							body.SetAtmElement "N2", GetAtmosphericGasQuantity(body, 100, 100)
					End Select

				''''''''''
				' Garden '
				''''''''''
				Case 6
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = GetHydrographicCoverage(1, 4)
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = GetAtmosphericPressure(body, 1)
					End If
					Select Case GetGURPSValue(body, "atmospheric_composition")
						''''''''''''''
						' Breathable '
						''''''''''''''
						Case 1
							BuildBreathableAtmosphere body

						''''''''''''
						' Marginal '
						''''''''''''
						Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11
							BuildMarginalAtmosphereGases body
					End Select

				''''''''''''''
				' Greenhouse '
				''''''''''''''
				Case 7
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = GetHydrographicCoverage(2, -7)
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = GetAtmosphericPressure(body, 100)
					End If
					Select Case GetGURPSValue(body, "atmospheric_composition")
						''''''''''''''''''''''''''''''''''''''''''''''
						' Suffocating, Lethally Toxic, and Corrosive '
						''''''''''''''''''''''''''''''''''''''''''''''
						Case 15
							body.SetAtmElement "N2", GetAtmosphericGasQuantity(body, 1, 5)
							body.SetAtmElement "C02", GetAtmosphericGasQuantity(body, 100, 100)
					End Select

				'''''''''''''
				' Chthonian '
				'''''''''''''
				Case 8
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = 0
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = Rand(1, 10) / 1000
					End If

			End Select

		'''''''''
		' Large '
		'''''''''
		Case 3
			Select Case GetGURPSValue(body, "terrestrial_type")
				'''''''
				' Ice '
				'''''''
				Case 1
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = GetHydrographicCoverage(2, -10)
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = GetAtmosphericPressure(body, 5)
					End If
					Select Case GetGURPSValue(body, "atmospheric_composition")
						''''''''''''''''''''''''''''''''
						' Suffocating and Highly Toxic '
						''''''''''''''''''''''''''''''''
						Case 14
							body.SetAtmElement "N2", GetAtmosphericGasQuantity(body, 10, 50)
							body.SetAtmElement "He", GetAtmosphericGasQuantity(body, 100, 100)
					End Select

				'''''''''''
				' Ammonia '
				'''''''''''
				Case 4
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = GetHydrographicCoverage(2, 0)
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = GetAtmosphericPressure(body, 5)
					End If
					Select Case GetGURPSValue(body, "atmospheric_composition")
						''''''''''''''''''''''''''''''''''''''''''''''
						' Suffocating, Lethally Toxic, and Corrosive '
						''''''''''''''''''''''''''''''''''''''''''''''
						Case 15
							body.SetAtmElement "NH3", GetAtmosphericGasQuantity(body, 10, 30)
							body.SetAtmElement "CH4", GetAtmosphericGasQuantity(body, 10, 30)
							body.SetAtmElement "He", GetAtmosphericGasQuantity(body, 100, 100)
					End Select

				'''''''''
				' Ocean '
				'''''''''
				Case 5
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = GetHydrographicCoverage(1, 6)
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = GetAtmosphericPressure(body, 5)
					End If
					Select Case GetGURPSValue(body, "atmospheric_composition")
						''''''''''''''''''''''''''''''''
						' Suffocating and Highly Toxic '
						''''''''''''''''''''''''''''''''
						Case 14
							body.SetAtmElement "N2", GetAtmosphericGasQuantity(body, 10, 50)
							body.SetAtmElement "He", GetAtmosphericGasQuantity(body, 100, 100)
					End Select

				''''''''''
				' Garden '
				''''''''''
				Case 6
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = GetHydrographicCoverage(1, 6)
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = GetAtmosphericPressure(body, 5)
					End If
					Select Case GetGURPSValue(body, "atmospheric_composition")
						''''''''''''''
						' Breathable '
						''''''''''''''
						Case 1
							BuildBreathableAtmosphere body

						''''''''''''
						' Marginal '
						''''''''''''
						Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11
							BuildMarginalAtmosphereGases body
					End Select

				''''''''''''''
				' Greenhouse '
				''''''''''''''
				Case 7
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = GetHydrographicCoverage(2, -7)
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = GetAtmosphericPressure(body, 500)
					End If
					Select Case GetGURPSValue(body, "atmospheric_composition")
						''''''''''''''''''''''''''''''''''''''''''''''
						' Suffocating, Lethally Toxic, and Corrosive '
						''''''''''''''''''''''''''''''''''''''''''''''
						Case 15
							body.SetAtmElement "N2", GetAtmosphericGasQuantity(body, 1, 5)
							body.SetAtmElement "C02", GetAtmosphericGasQuantity(body, 100, 100)
					End Select

				'''''''''''''
				' Chthonian '
				'''''''''''''
				Case 8
					If GetGURPSValue(sector, "hydrographic_coverage_flag") = 1 Then
						body.Water = 0
					End If
					body.ResetAtmElements()
					If GetGURPSValue(sector, "atmospheric_pressure_flag") = 1 Then
						body.Atmosphere = Rand(1, 10) / 1000
					End If

			End Select
	End Select

	'Sort the atmospheric gases
	If body.GetAtmElementCount > 0 Then
		SortAtmosphericGas(body)
	End If

End Sub


'------------------------------------------------------------------------------
'Build Marginal Atmosphere Gases
'------------------------------------------------------------------------------
Sub BuildMarginalAtmosphereGases(body)
	'Clear Atmospheric Gases
	body.ResetAtmElements()

	'Select a random inert gas, just for flavor
	If GetGURPSValue(body, "terrestrial_size") = 2 Then
		Select Case RollDice(1,5,0)
			Case 1
				body.SetAtmElement "He", GetAtmosphericGasQuantity(body, 0.1, 1)
			Case 2
				body.SetAtmElement "Ne", GetAtmosphericGasQuantity(body, 0.1, 1)
			Case 3
				body.SetAtmElement "Ar", GetAtmosphericGasQuantity(body, 0.1, 1)
			Case 4
				body.SetAtmElement "Kr", GetAtmosphericGasQuantity(body, 0.1, 1)
			Case 5
				body.SetAtmElement "Xe", GetAtmosphericGasQuantity(body, 0.1, 1)
		End Select
		body.SetAtmElement "02", GetAtmosphericGasQuantity(body, 18, 24)

	Else
		Select Case RollDice(1,5,0)
			Case 1
				body.SetAtmElement "He", GetAtmosphericGasQuantity(body, 1, 10)
			Case 2
				body.SetAtmElement "Ne", GetAtmosphericGasQuantity(body, 1, 10)
			Case 3
				body.SetAtmElement "Ar", GetAtmosphericGasQuantity(body, 1, 10)
			Case 4
				body.SetAtmElement "Kr", GetAtmosphericGasQuantity(body, 1, 10)
			Case 5
				body.SetAtmElement "Xe", GetAtmosphericGasQuantity(body, 1, 10)
		End Select
		body.SetAtmElement "02", GetAtmosphericGasQuantity(body, 25, 27)
	End If

	Select Case GetGURPSValue(body, "atmospheric_composition")
		''''''''''''''''''''''
		' Marginal: Chlorine '
		''''''''''''''''''''''
		Case 2
			body.SetAtmElement "Cl", GetAtmosphericGasQuantity(body, 0.003, 0.006)

		''''''''''''''''''''''
		' Marginal: Fluorine '
		''''''''''''''''''''''
		Case 3
			body.SetAtmElement "F", GetAtmosphericGasQuantity(body, 0.003, 0.006)

		''''''''''''''''''''''''''''''
		' Marginal: Sulfur Compounds '
		''''''''''''''''''''''''''''''
		Case 4
			Select Case RollDice(1,2,0)
				Case 1
					body.SetAtmElement "H2S", GetAtmosphericGasQuantity(body, 0.001, 0.025)
				Case 2
					body.SetAtmElement "SO2", GetAtmosphericGasQuantity(body, 0.000003, 0.000014)
			End Select

		''''''''''''''''''''''''''''''''
		' Marginal: Nitrogen Compounds '
		''''''''''''''''''''''''''''''''
		Case 5
			Select Case RollDice(1,3,0)
				Case 1
					body.SetAtmElement "NO", GetAtmosphericGasQuantity(body, 0.006, 0.015)
				Case 2
					body.SetAtmElement "N2O", GetAtmosphericGasQuantity(body, 8, 20)
				Case 3
					body.SetAtmElement "NO2", GetAtmosphericGasQuantity(body, 0.0005, 0.001)
			End Select

		''''''''''''''''''''''''''''
		' Marginal: Organic Toxins '
		''''''''''''''''''''''''''''
		Case 6
			body.SetAtmElement "Organic Toxins", GetAtmosphericGasQuantity(body, 0.01, 1)

		''''''''''''''''''''''''
		' Marginal: Low Oxygen '
		''''''''''''''''''''''''
		Case 7
			If GetGURPSValue(body, "terrestrial_size") = 2 Then
				body.SetAtmElement "02", GetAtmosphericGasQuantity(body, 15, 17)
			Else
				body.SetAtmElement "02", GetAtmosphericGasQuantity(body, 18, 24)
			End If

		''''''''''''''''''''''''
		' Marginal: Pollutants '
		''''''''''''''''''''''''
		Case 8
			Select Case RollDice(1,4,0)
				Case 1
					body.SetAtmElement "O3", GetAtmosphericGasQuantity(body, 0.000009, 0.004)
				Case 2
					body.SetAtmElement "CO", GetAtmosphericGasQuantity(body, 0.001, 0.01)
				Case 3
					body.SetAtmElement "SO2", GetAtmosphericGasQuantity(body, 0.000003, 0.000014)
				Case 4
					body.SetAtmElement "NO2", GetAtmosphericGasQuantity(body, 0.0005, 0.001)
			End Select

		'''''''''''''''''''''''''''''''''
		' Marginal: High Carbon Dioxide '
		'''''''''''''''''''''''''''''''''
		Case 9
			body.SetAtmElement "CO2", GetAtmosphericGasQuantity(body, 1, 10)

		'''''''''''''''''''''''''
		' Marginal: High Oxygen '
		'''''''''''''''''''''''''
		Case 10
			If GetGURPSValue(body, "terrestrial_size") = 2 Then
				body.SetAtmElement "02", GetAtmosphericGasQuantity(body, 25, 27)
			Else
				body.SetAtmElement "02", GetAtmosphericGasQuantity(body, 28, 30)
			End If

		'''''''''''''''''''''''''
		' Marginal: Inert Gases '
		'''''''''''''''''''''''''
		Case 11
			Select Case RollDice(1,6,0)
				Case 1
					body.SetAtmElement "He", GetAtmosphericGasQuantity(body, 0.05, 2)
				Case 2
					body.SetAtmElement "Ne", GetAtmosphericGasQuantity(body, 0.05, 2)
				Case 3
					body.SetAtmElement "Ar", GetAtmosphericGasQuantity(body, 0.05, 2)
				Case 4
					body.SetAtmElement "Kr", GetAtmosphericGasQuantity(body, 0.05, 2)
				Case 5
					body.SetAtmElement "Xe", GetAtmosphericGasQuantity(body, 0.05, 2)
				Case 6
					body.SetAtmElement "N2O", GetAtmosphericGasQuantity(body, 8, 20)
			End Select
	End Select

	'Add remainder as nitrogen
	body.SetAtmElement "N2", GetAtmosphericGasQuantity(body, 100, 100)

End Sub


'------------------------------------------------------------------------------
'Build Breathable Atmosphere
'------------------------------------------------------------------------------
Sub BuildBreathableAtmosphere(body)
	body.ResetAtmElements()    'Clear Atmospheric Gases

	'Select a random inert gas, just for flavor
	Select Case RollDice(1,5,0)
		Case 1
			body.SetAtmElement "He", GetAtmosphericGasQuantity(body, 0.1, 1)
		Case 2
			body.SetAtmElement "Ne", GetAtmosphericGasQuantity(body, 0.1, 1)
		Case 3
			body.SetAtmElement "Ar", GetAtmosphericGasQuantity(body, 0.1, 1)
		Case 4
			body.SetAtmElement "Kr", GetAtmosphericGasQuantity(body, 0.1, 1)
		Case 5
			body.SetAtmElement "Xe", GetAtmosphericGasQuantity(body, 0.1, 1)
	End Select
	body.SetAtmElement "02", GetAtmosphericGasQuantity(body, 18, 24)
	body.SetAtmElement "N2", GetAtmosphericGasQuantity(body, 100, 100)

	SetGURPSValue body, "atmospheric_composition", 1    'Breathable
End Sub


'------------------------------------------------------------------------------
'Build Marginal Atmosphere
'------------------------------------------------------------------------------
Sub BuildMarginalAtmosphere(body)
	Select Case RollDice(3,6,0)
		Case 3, 4
			If RollDice(1,100,0) <= 99 Then
				SetGURPSValue body, "atmospheric_composition", 2    'Marginal: Chlorine
			Else
				SetGURPSValue body, "atmospheric_composition", 3    'Marginal: Fluorine
			End If

		Case 5, 6
			SetGURPSValue body, "atmospheric_composition", 4        'Marginal: Sulfur Compounds

		Case 7
			SetGURPSValue body, "atmospheric_composition", 5        'Marginal: Nitrogen Compounds

		Case 8, 9
			SetGURPSValue body, "atmospheric_composition", 6        'Marginal: Organic Toxins

		Case 10, 11
			SetGURPSValue body, "atmospheric_composition", 7        'Marginal: Low Oxygen

		Case 12, 13
			SetGURPSValue body, "atmospheric_composition", 8        'Marginal: Pollutants

		Case 14
			SetGURPSValue body, "atmospheric_composition", 9        'Marginal: Pollutants

		Case 15, 16
			SetGURPSValue body, "atmospheric_composition", 1        'Marginal: High Oxygen

		Case 17, 18
			SetGURPSValue body, "atmospheric_composition", 11       'Marginal: Inert Gases
	End Select
End Sub


'------------------------------------------------------------------------------
'Build Atmospheric Density
'------------------------------------------------------------------------------
Sub BuildAtmosphereDensity(body)
	'Get the atmospheric density based on the atmospheric pressure
	If body.Atmosphere = 0 Then
		SetGURPSValue body, "atmospheric_density", 0    'None
	ElseIf body.Atmosphere <= 0.01 Then
		SetGURPSValue body, "atmospheric_density", 1    'Trace

	ElseIf body.Atmosphere <= 0.5 Then
		SetGURPSValue body, "atmospheric_density", 2    'Very Thin

	ElseIf body.Atmosphere <= 0.8 Then
		SetGURPSValue body, "atmospheric_density", 3    'Thin

	ElseIf body.Atmosphere <= 1.2 Then
		SetGURPSValue body, "atmospheric_density", 4    'Standard

	ElseIf body.Atmosphere <= 1.5 Then
		SetGURPSValue body, "atmospheric_density", 5    'Dense

	ElseIf body.Atmosphere <= 10 Then
		SetGURPSValue body, "atmospheric_density", 6    'Very Dense

	Else
		SetGURPSValue body, "atmospheric_density", 7    'Super Dense
	End If
End Sub


'------------------------------------------------------------------------------
'Build Climate Type
'------------------------------------------------------------------------------
Sub BuildClimateType(body)
	If body.Temp <= 244 Then
		SetGURPSValue body, "climate_type", 0    'Frozen

	ElseIf body.Temp <= 255 Then
		SetGURPSValue body, "climate_type", 1    'Very Cold

	ElseIf body.Temp <= 266 Then
		SetGURPSValue body, "climate_type", 2    'Cold

	ElseIf body.Temp <= 278 Then
		SetGURPSValue body, "climate_type", 3    'Chilly

	ElseIf body.Temp <= 289 Then
		SetGURPSValue body, "climate_type", 4    'Cool

	ElseIf body.Temp <= 300 Then
		SetGURPSValue body, "climate_type", 5    'Normal

	ElseIf body.Temp <= 311 Then
		SetGURPSValue body, "climate_type", 6    'Warm

	ElseIf body.Temp <= 322 Then
		SetGURPSValue body, "climate_type", 7    'Tropical

	ElseIf body.Temp <= 333 Then
		SetGURPSValue body, "climate_type", 8    'Hot

	ElseIf body.Temp <= 344 Then
		SetGURPSValue body, "climate_type", 9    'Very Hot

	Else
		SetGURPSValue body, "climate_type", 10   'Infernal
	End If
End Sub


'------------------------------------------------------------------------------
'Build Volcanic Activity
'------------------------------------------------------------------------------
Sub BuildVolcanicActivity(body)
	'Declare variables
	Dim tmp,moonCount,i    'as Int
	Dim stellarAge         'as Double

	'Determine the modifier based on the age and gravity of the planet
	'If the stellar age is 0 then the tmp will be 1 otherwise calculate
	'it out
	stellarAge = GetGURPSValue(body.GetRootBody, "stellar_age")
	If stellarAge = 0 Then
		tmp = 1
	Else
		tmp = Round(((body.Mass/(body.Radius/6378)^2) / stellarAge) * 40)
	End If

	'Make a roll and add it to the tmp
	tmp = tmp + RollDice(3,6,0)

	'Determine the modifier based on the number of moons
	'Loop through each child body
	moonCount = 0

	If body.ChildrenCount() > 0 Then
		For i = 1 To body.ChildrenCount()
			If body.GetChild(i - 1).TypeID = BODY_TYPE_TERRESTRIAL Then
				moonCount = moonCount + 1
			End If
		Next
	End If

	If moonCount = 1 Then
		tmp = tmp + 5
	ElseIf moonCount > 1 Then
		tmp = tmp + 10
	End If

	'Determine the modifier if the body is a gas giant moon
	If(body.HasParent()) Then
		If body.GetParentBody.TypeID = BODY_TYPE_GASGIANT Then
			If GetGURPSValue(body, "terrestrial_type") = 0 Then
				tmp = tmp + 60
			Else
				tmp = tmp + 5
			End If
		End If
	End If

	'Get the volcanic activity
	If tmp <= 16 Then
		SetGURPSValue body, "volcanic_activity", 0    'None

	ElseIf tmp <= 20 Then
		SetGURPSValue body, "volcanic_activity", 1    'Light

	ElseIf tmp <= 26 Then
		SetGURPSValue body, "volcanic_activity", 2    'Moderate

	ElseIf tmp <= 70 Then
		SetGURPSValue body, "volcanic_activity", 3    'Heavy

	Else
		SetGURPSValue body, "volcanic_activity", 4    'Extreme
	End If

	'Garden planets with heavy or extreme volcanic activity could have a marginal
	'atmosphere with sulfur compounds or pollutants
	If GetGURPSValue(body, "terrestrial_type") = 6 Then
		'Heavy
		If GetGURPSValue(body, "volcanic_activity") = 3 Then
			If RollDice(3,6,0) <= 8 Then
				If RollDice(1,100,0) < 50 Then
					SetGURPSValue body, "atmospheric_composition", 4
				Else
					SetGURPSValue body, "atmospheric_composition", 8
				End If
			End If

		'Extreme
		ElseIf GetGURPSValue(body, "volcanic_activity") = 4 Then
			If RollDice(3,6,0) <= 14 Then
				If RollDice(1,100,0) < 50 Then
					SetGURPSValue body, "atmospheric_composition", 4
				Else
					SetGURPSValue body, "atmospheric_composition", 8
				End If
			End If
		End If
	End If
End Sub


'------------------------------------------------------------------------------
'Build Tectonic Activity
'------------------------------------------------------------------------------
Sub BuildTectonicActivity(body)
	'Declare variables
	Dim tmp,moonCount,i    'as Int

	'Only check standard and large sized planets
	If GetGURPSValue(body, "terrestrial_size") >= 2 Then
		'Make a roll
		tmp = RollDice(3,6,0)

		'Determine the modifier for volcanic activity
		Select Case GetGURPSValue(body, "volcanic_activity")
			'None
			Case 0
				tmp = tmp - 8

			'Light
			Case 1
				tmp = tmp - 4

			'Heavy
			Case 3
				tmp = tmp + 4

			'Extreme
			Case 4
				tmp = tmp + 8
		End Select

		'Determine the modifier for hydrographic coverage
		If body.Water * 100 = 0 Then
			tmp = tmp - 4
		ElseIf body.Water * 100 < 50 Then
			tmp = tmp - 2
		End If

		'Determine the modifier based on the number of moons
		'Loop through each child body
		moonCount = 0
		If body.ChildrenCount() > 0 Then
			For i = 1 To body.ChildrenCount()
				If body.GetChild(i - 1).TypeID = BODY_TYPE_TERRESTRIAL Then
					moonCount = moonCount + 1
				End If
			Next
		End If

		If moonCount = 1 Then
			tmp = tmp + 2
		ElseIf moonCount > 1 Then
			tmp = tmp + 4
		End If

		'Get the tectonic activity
		If tmp <= 6 Then
			SetGURPSValue body, "tectonic_activity", 0    'None

		ElseIf tmp <= 10 Then
			SetGURPSValue body, "tectonic_activity", 1    'Light

		ElseIf tmp <= 14 Then
			SetGURPSValue body, "tectonic_activity", 2    'Moderate

		ElseIf tmp <= 18 Then
			SetGURPSValue body, "tectonic_activity", 3    'Heavy

		Else
			SetGURPSValue body, "tectonic_activity", 4    'Extreme
		End If

	'Tiny and Small planets
	Else
			SetGURPSValue body, "tectonic_activity", 0    'None
	End If
End Sub


'------------------------------------------------------------------------------
'Build Resource Value
'------------------------------------------------------------------------------
Sub BuildResourceValue(body)
	'Declare variables
	Dim tmp,i    'as Int

	'Make a roll
	tmp = RollDice(3,6,0)

	'Planets and asteroid belts have different tables for resources
	'Terrestrial
	If body.TypeID = BODY_TYPE_TERRESTRIAL Then
		'Modify for volcanic activity
		Select Case GetGURPSValue(body, "volcanic_activity")
			Case 0
				tmp = tmp - 2

			Case 1
				tmp = tmp - 1

			Case 3
				tmp = tmp + 1

			Case 4
				tmp = tmp + 2

		End Select

		'Get the resource value
		If tmp <= 2 Then
			SetGURPSValue body, "resource_value",  "-3"    'Scant

		ElseIf tmp <= 4 Then
			SetGURPSValue body, "resource_value",  "-2"    'Very Poor

		ElseIf tmp <= 7 Then
			SetGURPSValue body, "resource_value",  "-1"    'Poor

		ElseIf tmp <= 13 Then
			SetGURPSValue body, "resource_value",  "0"     'Average

		ElseIf tmp <= 16 Then
			SetGURPSValue body, "resource_value",  "1"     'Abundant

		ElseIf tmp <= 18 Then
			SetGURPSValue body, "resource_value",  "2"     'Very Abundant

		Else
			SetGURPSValue body, "resource_value",  "3"     'Rich
		End If

	'Asteroid Belt or Planetoid
	Else
		'Get the resource value
		If tmp = 3 Then
			SetGURPSValue body, "resource_value",  "-5"    'Worthless

		ElseIf tmp = 4 Then
			SetGURPSValue body, "resource_value",  "-4"    'Very Scant

		ElseIf tmp = 5 Then
			SetGURPSValue body, "resource_value",  "-3"    'Scant

		ElseIf tmp <= 7 Then
			SetGURPSValue body, "resource_value",  "-2"    'Very Poor

		ElseIf tmp <= 9 Then
			SetGURPSValue body, "resource_value",  "-1"    'Poor

		ElseIf tmp <= 11 Then
			SetGURPSValue body, "resource_value",  "0"    'Average

		ElseIf tmp <= 13 Then
			SetGURPSValue body, "resource_value",  "1"    'Abundant

		ElseIf tmp <= 15 Then
			SetGURPSValue body, "resource_value",  "2"    'Very Abundant

		ElseIf tmp = 16 Then
			SetGURPSValue body, "resource_value",  "3"    'Rich

		ElseIf tmp = 17 Then
			SetGURPSValue body, "resource_value",  "4"    'Very Rich

		Else
			SetGURPSValue body, "resource_value",  "5"    'Motherloade
		End If
	End If
End Sub


'------------------------------------------------------------------------------
'Build Habitability Value
'------------------------------------------------------------------------------
Sub BuildHabitabilityValue(body)
	'Declare variables
	Dim habitability,tmp    'as Int

	habitability = 0

	'No Atmosphere/Trace
	If GetGURPSValue(body, "atmospheric_density") <= "1" Then
		habitability = habitability + 0
	End If

	'Non Breathable Atmosphere (Very Thin or Above)
	If GetGURPSValue(body, "atmospheric_density") > "1" Then

		'Suffocating, Toxic, and Corrosive
		If GetGURPSValue(body, "atmospheric_composition") = "15" Then
			habitability = habitability - 2

		'Suffocating and Toxic only
		ElseIf GetGURPSValue(body, "atmospheric_composition") = "13" Or GetGURPSValue(body, "atmospheric_composition") = "14" Then
			habitability = habitability - 1

		'Suffocating only
		ElseIf GetGURPSValue(body, "atmospheric_composition") = "12" Then
			habitability = habitability + 0

		End If
	End If

	'Breathable/Marginal Atmosphere
	If GetGURPSValue(body, "atmospheric_composition") >= "1" And GetGURPSValue(body, "atmospheric_composition") <= "11" Then
		'Very Thin
		If GetGURPSValue(body, "atmospheric_density") = "2" Then
			habitability = habitability + 1

		'Thin
		ElseIf GetGURPSValue(body, "atmospheric_density") = "3" Then
			habitability = habitability + 2

		'Standard or Dense
		ElseIf GetGURPSValue(body, "atmospheric_density") = "4" Or GetGURPSValue(body, "atmospheric_density") = "5" Then
			habitability = habitability + 3

		'Very Dense or Super Dense
		ElseIf GetGURPSValue(body, "atmospheric_density") = "6" Or GetGURPSValue(body, "atmospheric_density") = "7" Then
			habitability = habitability + 1
		End If

		'Breathable Atmosphere Not Marginal
		If GetGURPSValue(body, "atmospheric_composition") = "1" Then
			habitability = habitability + 1
		End If

		'Frozen or Very Cold
		If GetGURPSValue(body, "climate_type") = "0" Or GetGURPSValue(body, "climate_type") = "1" Then
			habitability = habitability + 0

		'Cold
		ElseIf GetGURPSValue(body, "climate_type") = "2" Then
			habitability = habitability + 1

		'Chilly, Cool, Normal, Warm, or Tropical
		ElseIf GetGURPSValue(body, "climate_type") >= "3" And GetGURPSValue(body, "climate_type") <= "7" Then
			habitability = habitability + 2

		'Hot
		ElseIf GetGURPSValue(body, "climate_type") = "8" Then
			habitability = habitability + 1

		'Very Hot or Infernal
		Elseif GetGURPSValue(body, "climate_type") = "9" Or GetGURPSValue(body, "climate_type") = "10" Then
			habitability = habitability + 0
		End If
	End If

	'Hydrographic Coverage 0%
	If body.Water * 100 = 0 Then
		habitability = habitability + 0

	'Hydrographic Coverage 1% to 59%
	ElseIf body.Water * 100 <= 59 Then
		habitability = habitability + 1

	'Hydrographic Coverage 60% to 90%
	ElseIf body.Water * 100 <= 90 Then
		habitability = habitability + 2

	'Hydrographic Coverage 91% to 99%
	ElseIf body.Water * 100 <= 99 Then
		habitability = habitability + 1

	'Hydrographic Coverage 100%
	Else
		habitability = habitability + 0
	End If

	tmp = 0
	'Volcanic Activity
	If GetGURPSValue(body, "volcanic_activity") = "3" Then
		tmp = tmp + 1
	ElseIf GetGURPSValue(body, "volcanic_activity") = "4" Then
		tmp = tmp + 2
	End If

	'Tectonic Activity
	If GetGURPSValue(body, "tectonic_activity") = "3" Then
		tmp = tmp + 1
	ElseIf GetGURPSValue(body, "tectonic_activity") = "4" Then
		tmp = tmp + 2
	End If

	'Minimum of 2 if we have volcanic or tectonic activity
	If tmp <> 0 Then
		If tmp < 2 Then
			tmp = 2
		End If
	End If

	habitability = habitability - tmp

	'Lowest habitability value is -2
	If habitability < -2 Then
		habitability = -2
	End If

	'Highest habitability value is 8
	If habitability > 8 Then
		habitability = 8
	End If

	If GetGURPSValue(sector, "habitability_rating_flag") = 1 Then
		'Set the habitability value in AS
		If (habitability) <= 0 Then
			body.Habitability = HAB_INHOSPITABLE    'Inhospitable
		ElseIf (habitability) <= 3 Then
			body.Habitability = HAB_HABITABLE       'Habitable
		Else
			body.Habitability = HAB_HOSPITABLE      'Hospitable
		End If
	End If

	SetGURPSValue body, "habitability_value", habitability
End Sub


'------------------------------------------------------------------------------
'Get the Affinity Score
'------------------------------------------------------------------------------
Function GetAffinityScore(body)
	Dim affinity        'as int

	affinity = CInt(GetGURPSValue(body, "resource_value")) + CInt(GetGURPSValue(body, "habitability_value"))

	'Lowest affinity score is -5
	If affinity < -5 Then
		affinity = -5
	End If

	'Highest affinity score is 10
	If affinity > 10 Then
		affinity = 10
	End If

	GetAffinityScore = affinity
End Function


'------------------------------------------------------------------------------
'Get the the Atmospheric Pressure
'------------------------------------------------------------------------------
Function GetAtmosphericPressure(body, pressureFactor)
	'Declare variables
	Dim mass    'as Double

	mass = RollDice(3, 6, 0) / 10

	If Rand(1, 100) > 50 Then
		'Vary by up to 0.05 down
		mass = mass - Rand(0, 5) / 100
	Else
		'Vary by up to 0.05 up
		mass = mass + Rand(0, 5) / 100
	End If

	GetAtmosphericPressure = FormatNumber(mass * pressureFactor * (body.Mass/(body.Radius/6378)^2),2,,,0)
End Function

'------------------------------------------------------------------------------
'Get the Quantity of Atmospheric Gas for A Planet
'------------------------------------------------------------------------------
Function GetAtmosphericGasQuantity(body, gasMin, gasMax)
	'Declare variables
	Dim i, newGas, gasTotal    'as int

	gasMin = gasMin * 100
	gasMax = gasMax * 100

	'Loop through current gas to get the total
	For i = 1 To body.GetAtmElementCount
		gasTotal = gasTotal + body.GetAtmElementPercent(i - 1) * 100
	Next

	'Get random gas
	newGas = Rand(gasMin, gasMax)

	If newGas + gasTotal > 9800 Then
		newGas = 9800 - gasTotal
	End If

	'Return the gas quantity
	GetAtmosphericGasQuantity = newGas/100
End Function


'------------------------------------------------------------------------------
'Sort the Atmospheric Gas for A Planet
'------------------------------------------------------------------------------
Sub SortAtmosphericGas(body)
	'Declare variables
	Dim i, j                     'as int
	Dim gasValue(), gasName()    'as array

	'ReDim the array and put in an array to sort
	ReDim Preserve gasValue(body.GetAtmElementCount)
	ReDim Preserve gasName(body.GetAtmElementCount)

	'Loop through current gas to get the total
	For i = 1 To body.GetAtmElementCount
		gasValue(i - 1) = body.GetAtmElementPercent(i - 1)
		gasName(i - 1) = body.GetAtmElement(i - 1)
	Next

	'Sort the gases decending
	For i = LBound(gasValue) to UBound(gasValue)
		For j = LBound(gasValue) to UBound(gasValue)
			If j <> UBound(gasValue) Then
				If gasValue(j) < gasValue(j + 1) Then
						tmp = gasValue(j + 1)
						tmp2 = gasName(j +1 )
						gasValue(j + 1) = gasValue(j)
						gasName(j + 1) = gasName(j)
						gasValue(j) = tmp
						gasName(j) = tmp2
				 End If
			End If
		Next
	Next

	'Reset the current atmospheric gases
	body.ResetAtmElements()

	'Insert sorted gases
	For i = 0 To Ubound(gasValue)
		body.SetAtmElement gasName(i), gasValue(i)
	Next
End Sub


'------------------------------------------------------------------------------
'Get the Hydrographic Coverage for a Terrestrial Body
'------------------------------------------------------------------------------
Function GetHydrographicCoverage(diceCount, modifier)
	'Declare variables
	Dim baseHydro    'as Int

	baseHydro = RollDice(diceCount, 6, modifier) * 10

	If Rand(1, 100) > 50 Then
		'Vary by up to 5 down
		baseHydro = baseHydro - Rand(0, 500) / 100
	Else
		'Vary by up to 5 up
		baseHydro = baseHydro + Rand(0, 500) / 100
	End If

	If baseHydro > 100 Then
		baseHydro = 100
	End If

	If baseHydro < 0 Then
		baseHydro = 0
	End If

	GetHydrographicCoverage = FormatNumber(baseHydro/100,4,,,0)
End Function

'------------------------------------------------------------------------------
'Set the Stellar Age for All Stars in a System
'------------------------------------------------------------------------------
Sub SetStellarAge(body, stellarAge)
	'Declare variables
	Dim i    'as Int

	'loop through each child in a multiple star systems
	If body.TypeID = BODY_TYPE_MULT Or body.TypeID = BODY_TYPE_CLOSEMULT Then
		'Set the stellar age
		SetGURPSValue body, "stellar_age", stellarAge

		'Loop through each child
		If body.ChildrenCount() > 0 Then
			For i = 1 To body.ChildrenCount()
				SetStellarAge body.GetChild(i - 1), stellarAge
			Next
		End If
	End If

	'Set the stellar age for a star
	If body.TypeID = BODY_TYPE_STAR Then
		'Set the stellar age
		SetGURPSValue body, "stellar_age", stellarAge
		'Build Star
		BuildStarBody body
	End If

	'Set the stellar age for the terrestrial planet or gas giant
	'if they are the root body
	If Not body.HasParent() And body.TypeID = BODY_TYPE_TERRESTRIAL Or body.TypeID = BODY_TYPE_GASGIANT Then
		'Set the stellar age
		SetGURPSValue body, "stellar_age", stellarAge
	End If
End Sub


'------------------------------------------------------------------------------
'Randomly Select A Stellar Age
'------------------------------------------------------------------------------
Function GetStellarAge()
	'Declare variables
	Dim age, aBase, stepA, stepB    'as Double

	'Randomly pick an age
	Select Case RollDice(3,6,0)
		'Extreme Population I
		Case 3
			aBase = 0
			stepA = 0
			stepB = 0
		'Young Population I
		Case 4,5,6
			aBase = 0.1
			stepA = 0.3
			stepB = 0.05
		'Intermediate Population I
		Case 7,8,9,10
			aBase = 2
			stepA = 0.6
			stepB = 0.1
		'Old Population I
		Case 11,12,13,14
			aBase = 5.6
			stepA = 0.6
			stepB = 0.1
		'Intermediate Population II
		Case 15,16,17
			aBase = 8
			stepA = 0.6
			stepB = 0.1
		'Extreme Population II
		Case 18
			aBase = 10
			stepA = 0.6
			stepB = 0.1
	End Select

	'Calculate the age in billions of years
	age = (aBase + (RollDice(1,6,-1) * stepA) + (RollDice(1,6,-1) * stepB))

	'Return the stellar age and metallicity
	GetStellarAge = FormatNumber(age,1,,,0)
End Function


'------------------------------------------------------------------------------
'Display the Star Body
'------------------------------------------------------------------------------
Sub DisplayStarBody(body)
	'Declare variables
	Dim config    'as sdc

	body.SetField "Stellar Age", GetGURPSValue(body.GetRootBody, "stellar_age") & " Billion Years"

	config = CreateSystemDataConfig()
	config.AddField "Stellar Age", "custom field", "Stellar Age", TRUE
	config.AddField "Spectral Class", "spectral", "", FALSE
	config.AddField "Mass", "mass", "", TRUE
	config.AddField "Radius", "solradius", "", TRUE
	config.AddField "Luminosity", "luminosity", "", TRUE
	config.AddField "Political", "political", "", TRUE
	config.AddField "Planets", "children", "", TRUE

	body.SetSystemDataConfig(config)

	'Set the version that this body was created with
	SetGURPSValue body, "suite_version", GURPSSuiteVersion

	'Don't unload the body
	unloadBody = FALSE
End Sub


'------------------------------------------------------------------------------
'Display the Asteroid Belt Body
'------------------------------------------------------------------------------
Sub DisplayAsteroidBeltBody(body)
	'Declare variables
	Dim config    'as sdc

	body.SetField "Resource Value", RESOURCE_VALUE_ARRAY(GetGURPSValue(body, "resource_value")+5)
	body.SetField "Habitability Value", GetSign(GetGURPSValue(body, "habitability_value"))
	body.SetField "Affinity Score", GetSign(GetAffinityScore(body))

	config = CreateSystemDataConfig()
	config.AddField "Asteroid Belt", "", "", FALSE
	config.AddField "Distance", "distance", "", FALSE
	config.AddField "Belt Width", "radius", "", FALSE
	config.AddField "Resource Value", "custom field", "Resource Value", FALSE
	config.AddField "Habitability Value", "custom field", "Habitability Value", FALSE
	config.AddField "Affinity Score", "custom field", "Affinity Score", FALSE
	config.AddField "Political", "political", "", FALSE
	config.AddField "Population", "population", "", FALSE
	config.AddField "Tech Level", "custom field", "Tech Level", FALSE
	config.AddField "Contents", "children", "", TRUE

	body.SetSystemDataConfig(config)

	'Set the version that this body was created with
	SetGURPSValue body, "suite_version", GURPSSuiteVersion

	'Don't unload the body
	unloadBody = FALSE
End Sub


'------------------------------------------------------------------------------
'Display the Planetoid Body
'------------------------------------------------------------------------------
Sub DisplayPlanetoidBody(body)
	'Declare variables
	Dim config    'as sdc

	body.SetField "Climate", CLIMATE_TYPE_ARRAY(GetGURPSValue(body, "climate_type"))
	body.SetField "Resource Value", RESOURCE_VALUE_ARRAY(GetGURPSValue(body, "resource_value")+5)
	body.SetField "Habitability Value", GetSign(GetGURPSValue(body, "habitability_value"))
	body.SetField "Affinity Score", GetSign(GetAffinityScore(body))

	config = CreateSystemDataConfig()
	config.AddField "Planetoid", "", "", FALSE
	config.AddField "Distance", "distance", "", FALSE
	config.AddField "Radius", "radius", "", FALSE
	config.AddField "Gravity", "gravity", "", FALSE
	config.AddField "Orbit Period", "orbitperiod", "", FALSE
	config.AddField "Rotation", "rotation", "", FALSE
	config.AddField "Climate", "custom field", "Climate", FALSE
	config.AddField "Mean Temp", "temperature", "", FALSE
	config.AddField "Resource Value", "custom field", "Resource Value", FALSE
	config.AddField "Habitability Value", "custom field", "Habitability Value", FALSE
	config.AddField "Affinity Score", "custom field", "Affinity Score", FALSE
	config.AddField "", "habitability", "", TRUE
	config.AddField "Political", "political", "", FALSE
	config.AddField "Population", "population", "", FALSE
	config.AddField "Moons", "children", "", TRUE
	body.SetSystemDataConfig(config)

	'Set the version that this body was created with
	SetGURPSValue body, "suite_version", GURPSSuiteVersion

	'Don't unload the body
	unloadBody = FALSE
End Sub


'------------------------------------------------------------------------------
'Display the Gas Giant Body
'------------------------------------------------------------------------------
Sub DisplayGasGiantBody(body)
	'Declare variables
	Dim config    'as sdc

	body.SetField "Gas Giant Size", GAS_SIZE_ARRAY(GetGURPSValue(body, "gas_size"))

	config = CreateSystemDataConfig()
	config.AddField GAS_SIZE_ARRAY(GetGURPSValue(body, "gas_size")) & " Gas Giant", "", "", FALSE
	config.AddField "Distance", "distance", "", FALSE
	config.AddField "Radius", "radius", "", FALSE
	config.AddField "Gravity", "gravity", "", FALSE
	config.AddField "Orbit Period", "orbitperiod", "", FALSE
	config.AddField "Rotation", "rotation", "", FALSE
	config.AddField "Mean Temp", "temperature", "", FALSE
	config.AddField "Moons", "children", "", TRUE

	body.SetSystemDataConfig(config)

	'Set the version that this body was created with
	SetGURPSValue body, "suite_version", GURPSSuiteVersion

	'Don't unload the body
	unloadBody = FALSE
End Sub


'------------------------------------------------------------------------------
'Display the Terrestrial Body
'------------------------------------------------------------------------------
Sub DisplayTerrestrialBody(body)
	'Declare variables
	Dim config    'as sdc

	body.SetField "Terrestrial Type", TERRESTRIAL_TYPE_ARRAY(GetGURPSValue(body, "terrestrial_type"))
	body.SetField "Terrestrial Size", TERRESTRIAL_SIZE_ARRAY(GetGURPSValue(body, "terrestrial_size"))
	body.SetField "Atmospheric Composition", "." & ATMOSPHERIC_COMPOSITION_ARRAY(GetGURPSValue(body, "atmospheric_composition")) & "."
	body.SetField "Atmospheric Density", "." & ATMOSPHERIC_DENSITY_ARRAY(GetGURPSValue(body, "atmospheric_density")) & "."
	body.SetField "Atmosphere", ATMOSPHERIC_DENSITY_ARRAY(GetGURPSValue(body, "atmospheric_density")) & " (" & ATMOSPHERIC_COMPOSITION_ARRAY(GetGURPSValue(body, "atmospheric_composition")) & ")"
	body.AtmosphereNotes = body.GetField("Atmosphere")
	body.SetField "Climate", CLIMATE_TYPE_ARRAY(GetGURPSValue(body, "climate_type"))
	body.SetField "Volcanic Activity", GEOLOGICAL_ACTIVITY_ARRAY(GetGURPSValue(body, "volcanic_activity"))
	body.SetField "Tectonic Activity", GEOLOGICAL_ACTIVITY_ARRAY(GetGURPSValue(body, "tectonic_activity"))
	body.SetField "Resource Value", RESOURCE_VALUE_ARRAY(GetGURPSValue(body, "resource_value")+5)
	body.SetField "Habitability Value", GetSign(GetGURPSValue(body, "habitability_value"))
	body.SetField "Affinity Score", GetSign(GetAffinityScore(body))

	config = CreateSystemDataConfig()
	config.AddField TERRESTRIAL_SIZE_ARRAY(GetGURPSValue(body, "terrestrial_size")) & " (" & TERRESTRIAL_TYPE_ARRAY(GetGURPSValue(body, "terrestrial_type")) & ")", "", "", FALSE
	config.AddField "Distance", "distance", "", FALSE
	config.AddField "Radius", "radius", "", FALSE
	config.AddField "Gravity", "gravity", "", FALSE
	config.AddField "Orbit Period", "orbitperiod", "", FALSE
	config.AddField "Rotation", "rotation", "", FALSE
	config.AddField "Climate", "custom field", "Climate", FALSE
	config.AddField "Mean Temp", "temperature", "", FALSE
	config.AddField "Atmosphere", "atmosphere", "", FALSE
	config.AddField "Water / Ice Index", "water", "", FALSE
	config.AddField "Volcanic Activity", "custom field", "Volcanic Activity", FALSE
	config.AddField "Tectonic Activity", "custom field", "Tectonic Activity", FALSE
	config.AddField "Resource Value", "custom field", "Resource Value", FALSE
	config.AddField "Habitability Value", "custom field", "Habitability Value", FALSE
	config.AddField "Affinity Score", "custom field", "Affinity Score", FALSE
	config.AddField "", "habitability", "", TRUE
	config.AddField "Political", "political", "", FALSE
	config.AddField "Population", "population", "", FALSE
	'config.AddField "Tech Level", "custom field", "Tech Level", FALSE
	config.AddField "Moons", "children", "", TRUE

	body.SetSystemDataConfig(config)

	'Set the version that this body was created with
	SetGURPSValue body, "suite_version", GURPSSuiteVersion

	'Don't unload the body
	unloadBody = FALSE
End Sub


'------------------------------------------------------------------------------
'Add a + to the Front of a Positive Intiger
'------------------------------------------------------------------------------
Function GetSign(inVal)
	If Left(inVal, 1) <> "-" Then
		GetSign = "+" & inVal
	Else
		GetSign = inVal
	End If
End Function


'------------------------------------------------------------------------------
'Get a Random Number Between the Min and Max
'------------------------------------------------------------------------------
Function Rand(minRand, maxRand)
	'Initialize the randomizer
	Randomize

	'Declare variables
	Dim minDecimalPlaces, maxDecimalPlaces, result    'as Double
	Dim multipier                                     'as Int

	If InStrRev(minRand, ".") > 0 Then
			minDecimalPlaces = Len(minRand) - InStrRev(minRand, ".")
	Else
			minDecimalPlaces = 0
	End If

	If InStrRev(maxRand, ".") > 0 Then
			maxDecimalPlaces = Len(maxRand) - InStrRev(maxRand, ".")
	Else
			maxDecimalPlaces = 0
	End If

	If minDecimalPlaces >= maxDecimalPlaces Then
		multipier = 10^minDecimalPlaces
	Else
		multipier = 10^maxDecimalPlaces
	End If

	minRand = minRand * multipier
	maxRand = maxRand * multipier

	result = Int((maxRand-minRand+1)*Rnd+minRand)

	Rand = result/multipier
End Function


'------------------------------------------------------------------------------
'Get Primary Star
'------------------------------------------------------------------------------
Function GetPrimaryStar(body, primaryStar)
	'Declare variables
	Dim i    'as Int

	'loop through each child in a multiple star systems
	If body.TypeID = BODY_TYPE_MULT Or body.TypeID = BODY_TYPE_CLOSEMULT Then
		'Loop through each child
		If body.ChildrenCount() > 0 Then
			For i = 1 To body.ChildrenCount()
				primaryStar = GetPrimaryStar(body.GetChild(i - 1), primaryStar)
			Next
		End If
	End If

	If body.TypeID = BODY_TYPE_STAR Then
		If Not IsNull(primaryStar) Then
			If body.Mass > primaryStar.Mass Then
				primaryStar = body
			End If
		Else
			primaryStar = body
		End If
	End If

	GetPrimaryStar = primaryStar
End Function


'------------------------------------------------------------------------------
'Set the Default for the Configuration
'------------------------------------------------------------------------------
Sub SetDefaultConfig()
	'Set default values
	If GetGURPSValue(sector, "atmospheric_pressure_flag") = "" Then
		SetGURPSValue sector, "atmospheric_pressure_flag", 0
	End If

	If GetGURPSValue(sector, "hydrographic_coverage_flag") = "" Then
		SetGURPSValue sector, "hydrographic_coverage_flag", 0
	End If

	If GetGURPSValue(sector, "habitability_rating_flag") = "" Then
		SetGURPSValue sector, "habitability_rating_flag", 0
	End If
End Sub

'------------------------------------------------------------------------------
'Set the GURPS Value
'------------------------------------------------------------------------------
Sub SetGURPSValue(body, GURPSField, GURPSVal)
	'Set the bodies field
	body.SetField "gurps_" & GURPSField & "_value", GURPSVal

	'Don't unload the body
	unloadBody = FALSE
End Sub


'------------------------------------------------------------------------------
'Get the GURPS Value
'------------------------------------------------------------------------------
Function GetGURPSValue(body, GURPSField)
	'Return the value
	GetGURPSValue = body.GetField("gurps_" & GURPSField & "_value")
End Function


'------------------------------------------------------------------------------
'Set the GURPS Lock
'------------------------------------------------------------------------------
Sub SetGURPSLock(body, GURPSField, GURPSVal)
	'Set the bodies field
	body.SetField "gurps_" & GURPSField & "_lock", GURPSVal

	'Don't unload the body
	unloadBody = FALSE
End Sub


'------------------------------------------------------------------------------
'Get the GURPS Lock
'------------------------------------------------------------------------------
Function GetGURPSLock(body, GURPSField)
	'Return the value
	If body.GetField("gurps_" & GURPSField & "_lock") = "-1" Then
		GetGURPSLock = "-1"    'TRUE
	Else
		GetGURPSLock = "0"     'FALSE
	End If
End Function


'------------------------------------------------------------------------------
'Include Files
'------------------------------------------------------------------------------
Sub Include(FilePath)
	Dim FSO, s, TS
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set TS = FSO.OpenTextFile(PluginDirectory() & "\" & FilePath, 1, FALSE)
	s = TS.ReadAll
	TS.Close
	Set TS = Nothing
	Set FSO = Nothing
	ExecuteGlobal s
End Sub
