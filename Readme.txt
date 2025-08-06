Battle Calculator for FE8. Only handles damage, but does not include Poison damage (which I probably could add with little effort) or Pure Water (which would be more annoying without class modules). I believe it handles everything else, though haven't tested it completely. And I hoping any users would point out any miscalculations so I can fix them.

Healing is treated as negative damage so there will likely be some bugs.
Probabilities are mutually exclusive and are defined as follows:
	Death	- chance that any player units die or any enemy units die in battles where they are marked to live (enemyMustLive).
	Fail	- chance that all player units survive and any enemy units survive when they are marked for death (EnemyMustDie).
	Success - chance that all player units survive and all enemy units survive when they are marked to live and die when they are marked for death.
	

1.	Download the BattleCalculator.xlsm file, open and click "Enable Content". Will most likely have to check "Unblock" in file properties to enable macros.

2.	Make sure Microsoft Scripting Runtime in checked as I frequently use the Scripting.Dictionary object.
		In VBA editor (Alt+F11 if not) -> Tools -> References, check box relevant box (might have to scroll down a bit).

3.	Enter data in the UnitData table (aka the full green one with PascalCase headers).
		Name is just a string identifier.
		Class MUST BE in the ClassData table. There is a dropdown list.
		Level to Con are all integers.
		Srank SHOULD BE whatever type (Category in WeaponData). Doesn't break if entered incorrectly but may lead to inaccurate calculations.
		HoplonGuard/FiliShield should be TRUE if unit holds the item, otherwise leave blank (or enter FALSE).
		IsPlayer should be TRUE if unit is player unit, otherwise leave blank (or enter FALSE).
		EnemyMustDie determines where states and their respective probabilities are tallied into (success or fail).

4.	Enter data in the CombatData table (aka the one with 8 green columns and has camelCase headers).
		atk/def is just attacker/defender.
		atkName to defTer have dropdown lists.
		canCounter is TRUE if the defender counters.
			This is not validated, the subroutine will compute a Purge counterattack against a Javelin if you tell it to.
		enemyMustLive is TRUE if you want any states where the enemy (if there is one) is dead at the end of this battle to be tallied to death prob.

5.	Additional Notes regarding the CombatData table.
5a.	When using healing staves: 
		atkName/defName should be allied but there is currently no validation programmed for this
		canCounter should be blank (or set to FALSE)
		defWpn should be blank (or set to "Unequipped")
5b.	When using healing items or terrain healing:
		atkName/defName should be the same but there is currently no validation programmed for this
		canCounter should be blank (or set to FALSE)
		defWpn should be blank (or set to "Unequipped")

6.	Regarding UnitData table input validation:
		Quite lacking.
		There are dropdown lists and some basic validation with regards to inputs but thats really about it.
		You can set a unit to be a Sniper with SRank "Sword", 30 Def, 25 Con.

7.	Regarding CombatData table input validation:
		Even more lacking.
		You can make a player Bishop heal an enemy Monster and it will compute as if it's a critical heal (3x heal).
		Leaving Wpn or Ter blank will result in "Unequipped" or "-" respectively.
		Only rows with non-empty atkName and defName will be processed, everything else is ignored.


Can also move both input tables (UnitData and CombatData around), the writes happen to Cells A1:B3 on the first page, and then the "Results" sheet.
