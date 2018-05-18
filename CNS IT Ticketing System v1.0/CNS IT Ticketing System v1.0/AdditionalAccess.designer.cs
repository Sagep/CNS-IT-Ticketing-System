﻿#pragma warning disable 1591
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace CNS_IT_Ticketing_System_v1._0
{
	using System.Data.Linq;
	using System.Data.Linq.Mapping;
	using System.Data;
	using System.Collections.Generic;
	using System.Reflection;
	using System.Linq;
	using System.Linq.Expressions;
	using System.ComponentModel;
	using System;
	
	
	[global::System.Data.Linq.Mapping.DatabaseAttribute(Name="ITTickets")]
	public partial class AdditionalAccessDataContext : System.Data.Linq.DataContext
	{
		
		private static System.Data.Linq.Mapping.MappingSource mappingSource = new AttributeMappingSource();
		
    #region Extensibility Method Definitions
    partial void OnCreated();
    partial void InsertAdditionalAccess(AdditionalAccess instance);
    partial void UpdateAdditionalAccess(AdditionalAccess instance);
    partial void DeleteAdditionalAccess(AdditionalAccess instance);
    #endregion
		
		public AdditionalAccessDataContext() : 
				base(global::CNS_IT_Ticketing_System_v1._0.Properties.Settings.Default.ITTicketsConnectionString3, mappingSource)
		{
			OnCreated();
		}
		
		public AdditionalAccessDataContext(string connection) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public AdditionalAccessDataContext(System.Data.IDbConnection connection) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public AdditionalAccessDataContext(string connection, System.Data.Linq.Mapping.MappingSource mappingSource) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public AdditionalAccessDataContext(System.Data.IDbConnection connection, System.Data.Linq.Mapping.MappingSource mappingSource) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public System.Data.Linq.Table<AdditionalAccess> AdditionalAccesses
		{
			get
			{
				return this.GetTable<AdditionalAccess>();
			}
		}
	}
	
	[global::System.Data.Linq.Mapping.TableAttribute(Name="dbo.AdditionalAccess")]
	public partial class AdditionalAccess : INotifyPropertyChanging, INotifyPropertyChanged
	{
		
		private static PropertyChangingEventArgs emptyChangingEventArgs = new PropertyChangingEventArgs(String.Empty);
		
		private int _GuestID;
		
		private string _GuestName;
		
		private string _GuestAccess;
		
    #region Extensibility Method Definitions
    partial void OnLoaded();
    partial void OnValidate(System.Data.Linq.ChangeAction action);
    partial void OnCreated();
    partial void OnGuestIDChanging(int value);
    partial void OnGuestIDChanged();
    partial void OnGuestNameChanging(string value);
    partial void OnGuestNameChanged();
    partial void OnGuestAccessChanging(string value);
    partial void OnGuestAccessChanged();
    #endregion
		
		public AdditionalAccess()
		{
			OnCreated();
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_GuestID", DbType="Int NOT NULL", IsPrimaryKey=true)]
		public int GuestID
		{
			get
			{
				return this._GuestID;
			}
			set
			{
				if ((this._GuestID != value))
				{
					this.OnGuestIDChanging(value);
					this.SendPropertyChanging();
					this._GuestID = value;
					this.SendPropertyChanged("GuestID");
					this.OnGuestIDChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_GuestName", DbType="Text", UpdateCheck=UpdateCheck.Never)]
		public string GuestName
		{
			get
			{
				return this._GuestName;
			}
			set
			{
				if ((this._GuestName != value))
				{
					this.OnGuestNameChanging(value);
					this.SendPropertyChanging();
					this._GuestName = value;
					this.SendPropertyChanged("GuestName");
					this.OnGuestNameChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_GuestAccess", DbType="Text", UpdateCheck=UpdateCheck.Never)]
		public string GuestAccess
		{
			get
			{
				return this._GuestAccess;
			}
			set
			{
				if ((this._GuestAccess != value))
				{
					this.OnGuestAccessChanging(value);
					this.SendPropertyChanging();
					this._GuestAccess = value;
					this.SendPropertyChanged("GuestAccess");
					this.OnGuestAccessChanged();
				}
			}
		}
		
		public event PropertyChangingEventHandler PropertyChanging;
		
		public event PropertyChangedEventHandler PropertyChanged;
		
		protected virtual void SendPropertyChanging()
		{
			if ((this.PropertyChanging != null))
			{
				this.PropertyChanging(this, emptyChangingEventArgs);
			}
		}
		
		protected virtual void SendPropertyChanged(String propertyName)
		{
			if ((this.PropertyChanged != null))
			{
				this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
			}
		}
	}
}
#pragma warning restore 1591
