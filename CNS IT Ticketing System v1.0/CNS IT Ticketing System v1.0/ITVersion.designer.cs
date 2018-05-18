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
	public partial class ITVersionDataContext : System.Data.Linq.DataContext
	{
		
		private static System.Data.Linq.Mapping.MappingSource mappingSource = new AttributeMappingSource();
		
    #region Extensibility Method Definitions
    partial void OnCreated();
    partial void InsertITTVersion(ITTVersion instance);
    partial void UpdateITTVersion(ITTVersion instance);
    partial void DeleteITTVersion(ITTVersion instance);
    #endregion
		
		public ITVersionDataContext() : 
				base(global::CNS_IT_Ticketing_System_v1._0.Properties.Settings.Default.ITTicketsConnectionString2, mappingSource)
		{
			OnCreated();
		}
		
		public ITVersionDataContext(string connection) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public ITVersionDataContext(System.Data.IDbConnection connection) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public ITVersionDataContext(string connection, System.Data.Linq.Mapping.MappingSource mappingSource) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public ITVersionDataContext(System.Data.IDbConnection connection, System.Data.Linq.Mapping.MappingSource mappingSource) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public System.Data.Linq.Table<ITTVersion> ITTVersions
		{
			get
			{
				return this.GetTable<ITTVersion>();
			}
		}
	}
	
	[global::System.Data.Linq.Mapping.TableAttribute(Name="dbo.ITTVersion")]
	public partial class ITTVersion : INotifyPropertyChanging, INotifyPropertyChanged
	{
		
		private static PropertyChangingEventArgs emptyChangingEventArgs = new PropertyChangingEventArgs(String.Empty);
		
		private double _Id;
		
		private string _Messagetoshow;
		
    #region Extensibility Method Definitions
    partial void OnLoaded();
    partial void OnValidate(System.Data.Linq.ChangeAction action);
    partial void OnCreated();
    partial void OnIdChanging(double value);
    partial void OnIdChanged();
    partial void OnMessagetoshowChanging(string value);
    partial void OnMessagetoshowChanged();
    #endregion
		
		public ITTVersion()
		{
			OnCreated();
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Id", DbType="Float NOT NULL", IsPrimaryKey=true)]
		public double Id
		{
			get
			{
				return this._Id;
			}
			set
			{
				if ((this._Id != value))
				{
					this.OnIdChanging(value);
					this.SendPropertyChanging();
					this._Id = value;
					this.SendPropertyChanged("Id");
					this.OnIdChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Messagetoshow", DbType="Text", UpdateCheck=UpdateCheck.Never)]
		public string Messagetoshow
		{
			get
			{
				return this._Messagetoshow;
			}
			set
			{
				if ((this._Messagetoshow != value))
				{
					this.OnMessagetoshowChanging(value);
					this.SendPropertyChanging();
					this._Messagetoshow = value;
					this.SendPropertyChanged("Messagetoshow");
					this.OnMessagetoshowChanged();
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
