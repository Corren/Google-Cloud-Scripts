UPDATE [ADEPTO].[dbo].[Equipment]
   SET UDdate1 = LocatedOn, DeptOwnerKey = 1512, [Notes] = ISNULL(CAST(Notes AS VARCHAR(MAX)), '') 
   + '[Updated on 1/4/2018. Facility/Department now set to LEGACY/OLD and the Date Sold to Date Inactive. This item was sold to [' + ISNULL(CAST([LastLocation] AS VARCHAR(MAX)), '')+
   '] on [' + ISNULL(CAST(LocatedOn AS VARCHAR(MAX)), '')
	+   '] and the location was never updated.]'
   WHERE UDdate1 IS NULL
     AND [EquipmentStatKey] = 5
  and [DeptOwnerKey] = 2
  and [Inactive] = 1



  UPDATE [ADEPTO].[dbo].[Equipment]
   SET [DeptCharg2Key] = 1512
   WHERE [EquipmentStatKey] = 5
  and [DeptOwnerKey] = 1512
  and [Inactive] = 1