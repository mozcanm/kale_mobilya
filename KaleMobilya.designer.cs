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

namespace Kale_Mobilya
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
	
	
	[global::System.Data.Linq.Mapping.DatabaseAttribute(Name="KaleMobilya")]
	public partial class KaleMobilyaDataContext : System.Data.Linq.DataContext
	{
		
		private static System.Data.Linq.Mapping.MappingSource mappingSource = new AttributeMappingSource();
		
    #region Extensibility Method Definitions
    partial void OnCreated();
    partial void InsertCari(Cari instance);
    partial void UpdateCari(Cari instance);
    partial void DeleteCari(Cari instance);
    partial void InsertDurum(Durum instance);
    partial void UpdateDurum(Durum instance);
    partial void DeleteDurum(Durum instance);
    partial void InsertKisi(Kisi instance);
    partial void UpdateKisi(Kisi instance);
    partial void DeleteKisi(Kisi instance);
    partial void InsertSenet(Senet instance);
    partial void UpdateSenet(Senet instance);
    partial void DeleteSenet(Senet instance);
    #endregion
		
		public KaleMobilyaDataContext() : 
				base(global::Kale_Mobilya.Properties.Settings.Default.KaleMobilyaConnectionString1, mappingSource)
		{
			OnCreated();
		}
		
		public KaleMobilyaDataContext(string connection) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public KaleMobilyaDataContext(System.Data.IDbConnection connection) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public KaleMobilyaDataContext(string connection, System.Data.Linq.Mapping.MappingSource mappingSource) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public KaleMobilyaDataContext(System.Data.IDbConnection connection, System.Data.Linq.Mapping.MappingSource mappingSource) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public System.Data.Linq.Table<Cari> Caris
		{
			get
			{
				return this.GetTable<Cari>();
			}
		}
		
		public System.Data.Linq.Table<Durum> Durums
		{
			get
			{
				return this.GetTable<Durum>();
			}
		}
		
		public System.Data.Linq.Table<Kisi> Kisis
		{
			get
			{
				return this.GetTable<Kisi>();
			}
		}
		
		public System.Data.Linq.Table<Senet> Senets
		{
			get
			{
				return this.GetTable<Senet>();
			}
		}
	}
	
	[global::System.Data.Linq.Mapping.TableAttribute(Name="dbo.Cari")]
	public partial class Cari : INotifyPropertyChanging, INotifyPropertyChanged
	{
		
		private static PropertyChangingEventArgs emptyChangingEventArgs = new PropertyChangingEventArgs(String.Empty);
		
		private int _CariID;
		
		private int _KisiID;
		
		private System.Nullable<byte> _DurumID;
		
		private System.Nullable<decimal> _Tutar;
		
		private System.Nullable<System.DateTime> _Tarih;
		
		private string _Aciklama;
		
		private EntityRef<Durum> _Durum;
		
		private EntityRef<Kisi> _Kisi;
		
    #region Extensibility Method Definitions
    partial void OnLoaded();
    partial void OnValidate(System.Data.Linq.ChangeAction action);
    partial void OnCreated();
    partial void OnCariIDChanging(int value);
    partial void OnCariIDChanged();
    partial void OnKisiIDChanging(int value);
    partial void OnKisiIDChanged();
    partial void OnDurumIDChanging(System.Nullable<byte> value);
    partial void OnDurumIDChanged();
    partial void OnTutarChanging(System.Nullable<decimal> value);
    partial void OnTutarChanged();
    partial void OnTarihChanging(System.Nullable<System.DateTime> value);
    partial void OnTarihChanged();
    partial void OnAciklamaChanging(string value);
    partial void OnAciklamaChanged();
    #endregion
		
		public Cari()
		{
			this._Durum = default(EntityRef<Durum>);
			this._Kisi = default(EntityRef<Kisi>);
			OnCreated();
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_CariID", AutoSync=AutoSync.OnInsert, DbType="Int NOT NULL IDENTITY", IsPrimaryKey=true, IsDbGenerated=true)]
		public int CariID
		{
			get
			{
				return this._CariID;
			}
			set
			{
				if ((this._CariID != value))
				{
					this.OnCariIDChanging(value);
					this.SendPropertyChanging();
					this._CariID = value;
					this.SendPropertyChanged("CariID");
					this.OnCariIDChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_KisiID", DbType="Int NOT NULL")]
		public int KisiID
		{
			get
			{
				return this._KisiID;
			}
			set
			{
				if ((this._KisiID != value))
				{
					if (this._Kisi.HasLoadedOrAssignedValue)
					{
						throw new System.Data.Linq.ForeignKeyReferenceAlreadyHasValueException();
					}
					this.OnKisiIDChanging(value);
					this.SendPropertyChanging();
					this._KisiID = value;
					this.SendPropertyChanged("KisiID");
					this.OnKisiIDChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_DurumID", DbType="TinyInt")]
		public System.Nullable<byte> DurumID
		{
			get
			{
				return this._DurumID;
			}
			set
			{
				if ((this._DurumID != value))
				{
					if (this._Durum.HasLoadedOrAssignedValue)
					{
						throw new System.Data.Linq.ForeignKeyReferenceAlreadyHasValueException();
					}
					this.OnDurumIDChanging(value);
					this.SendPropertyChanging();
					this._DurumID = value;
					this.SendPropertyChanged("DurumID");
					this.OnDurumIDChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Tutar", DbType="Decimal(10,2)")]
		public System.Nullable<decimal> Tutar
		{
			get
			{
				return this._Tutar;
			}
			set
			{
				if ((this._Tutar != value))
				{
					this.OnTutarChanging(value);
					this.SendPropertyChanging();
					this._Tutar = value;
					this.SendPropertyChanged("Tutar");
					this.OnTutarChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Tarih", DbType="Date")]
		public System.Nullable<System.DateTime> Tarih
		{
			get
			{
				return this._Tarih;
			}
			set
			{
				if ((this._Tarih != value))
				{
					this.OnTarihChanging(value);
					this.SendPropertyChanging();
					this._Tarih = value;
					this.SendPropertyChanged("Tarih");
					this.OnTarihChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Aciklama", DbType="NVarChar(500)")]
		public string Aciklama
		{
			get
			{
				return this._Aciklama;
			}
			set
			{
				if ((this._Aciklama != value))
				{
					this.OnAciklamaChanging(value);
					this.SendPropertyChanging();
					this._Aciklama = value;
					this.SendPropertyChanged("Aciklama");
					this.OnAciklamaChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.AssociationAttribute(Name="Durum_Cari", Storage="_Durum", ThisKey="DurumID", OtherKey="DurumID", IsForeignKey=true)]
		public Durum Durum
		{
			get
			{
				return this._Durum.Entity;
			}
			set
			{
				Durum previousValue = this._Durum.Entity;
				if (((previousValue != value) 
							|| (this._Durum.HasLoadedOrAssignedValue == false)))
				{
					this.SendPropertyChanging();
					if ((previousValue != null))
					{
						this._Durum.Entity = null;
						previousValue.Caris.Remove(this);
					}
					this._Durum.Entity = value;
					if ((value != null))
					{
						value.Caris.Add(this);
						this._DurumID = value.DurumID;
					}
					else
					{
						this._DurumID = default(Nullable<byte>);
					}
					this.SendPropertyChanged("Durum");
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.AssociationAttribute(Name="Kisi_Cari", Storage="_Kisi", ThisKey="KisiID", OtherKey="KisiID", IsForeignKey=true)]
		public Kisi Kisi
		{
			get
			{
				return this._Kisi.Entity;
			}
			set
			{
				Kisi previousValue = this._Kisi.Entity;
				if (((previousValue != value) 
							|| (this._Kisi.HasLoadedOrAssignedValue == false)))
				{
					this.SendPropertyChanging();
					if ((previousValue != null))
					{
						this._Kisi.Entity = null;
						previousValue.Caris.Remove(this);
					}
					this._Kisi.Entity = value;
					if ((value != null))
					{
						value.Caris.Add(this);
						this._KisiID = value.KisiID;
					}
					else
					{
						this._KisiID = default(int);
					}
					this.SendPropertyChanged("Kisi");
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
	
	[global::System.Data.Linq.Mapping.TableAttribute(Name="dbo.Durum")]
	public partial class Durum : INotifyPropertyChanging, INotifyPropertyChanged
	{
		
		private static PropertyChangingEventArgs emptyChangingEventArgs = new PropertyChangingEventArgs(String.Empty);
		
		private byte _DurumID;
		
		private string _Durumlar;
		
		private EntitySet<Cari> _Caris;
		
    #region Extensibility Method Definitions
    partial void OnLoaded();
    partial void OnValidate(System.Data.Linq.ChangeAction action);
    partial void OnCreated();
    partial void OnDurumIDChanging(byte value);
    partial void OnDurumIDChanged();
    partial void OnDurumlarChanging(string value);
    partial void OnDurumlarChanged();
    #endregion
		
		public Durum()
		{
			this._Caris = new EntitySet<Cari>(new Action<Cari>(this.attach_Caris), new Action<Cari>(this.detach_Caris));
			OnCreated();
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_DurumID", DbType="TinyInt NOT NULL", IsPrimaryKey=true)]
		public byte DurumID
		{
			get
			{
				return this._DurumID;
			}
			set
			{
				if ((this._DurumID != value))
				{
					this.OnDurumIDChanging(value);
					this.SendPropertyChanging();
					this._DurumID = value;
					this.SendPropertyChanged("DurumID");
					this.OnDurumIDChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Durumlar", DbType="NVarChar(7) NOT NULL", CanBeNull=false)]
		public string Durumlar
		{
			get
			{
				return this._Durumlar;
			}
			set
			{
				if ((this._Durumlar != value))
				{
					this.OnDurumlarChanging(value);
					this.SendPropertyChanging();
					this._Durumlar = value;
					this.SendPropertyChanged("Durumlar");
					this.OnDurumlarChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.AssociationAttribute(Name="Durum_Cari", Storage="_Caris", ThisKey="DurumID", OtherKey="DurumID")]
		public EntitySet<Cari> Caris
		{
			get
			{
				return this._Caris;
			}
			set
			{
				this._Caris.Assign(value);
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
		
		private void attach_Caris(Cari entity)
		{
			this.SendPropertyChanging();
			entity.Durum = this;
		}
		
		private void detach_Caris(Cari entity)
		{
			this.SendPropertyChanging();
			entity.Durum = null;
		}
	}
	
	[global::System.Data.Linq.Mapping.TableAttribute(Name="dbo.Kisi")]
	public partial class Kisi : INotifyPropertyChanging, INotifyPropertyChanged
	{
		
		private static PropertyChangingEventArgs emptyChangingEventArgs = new PropertyChangingEventArgs(String.Empty);
		
		private int _KisiID;
		
		private string _Ad;
		
		private string _Firma;
		
		private string _Tel1;
		
		private string _Tel2;
		
		private string _Adres;
		
		private System.Nullable<bool> _Karaliste;
		
		private EntitySet<Cari> _Caris;
		
		private EntitySet<Senet> _Senets;
		
    #region Extensibility Method Definitions
    partial void OnLoaded();
    partial void OnValidate(System.Data.Linq.ChangeAction action);
    partial void OnCreated();
    partial void OnKisiIDChanging(int value);
    partial void OnKisiIDChanged();
    partial void OnAdChanging(string value);
    partial void OnAdChanged();
    partial void OnFirmaChanging(string value);
    partial void OnFirmaChanged();
    partial void OnTel1Changing(string value);
    partial void OnTel1Changed();
    partial void OnTel2Changing(string value);
    partial void OnTel2Changed();
    partial void OnAdresChanging(string value);
    partial void OnAdresChanged();
    partial void OnKaralisteChanging(System.Nullable<bool> value);
    partial void OnKaralisteChanged();
    #endregion
		
		public Kisi()
		{
			this._Caris = new EntitySet<Cari>(new Action<Cari>(this.attach_Caris), new Action<Cari>(this.detach_Caris));
			this._Senets = new EntitySet<Senet>(new Action<Senet>(this.attach_Senets), new Action<Senet>(this.detach_Senets));
			OnCreated();
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_KisiID", AutoSync=AutoSync.OnInsert, DbType="Int NOT NULL IDENTITY", IsPrimaryKey=true, IsDbGenerated=true)]
		public int KisiID
		{
			get
			{
				return this._KisiID;
			}
			set
			{
				if ((this._KisiID != value))
				{
					this.OnKisiIDChanging(value);
					this.SendPropertyChanging();
					this._KisiID = value;
					this.SendPropertyChanged("KisiID");
					this.OnKisiIDChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Ad", DbType="NVarChar(31)")]
		public string Ad
		{
			get
			{
				return this._Ad;
			}
			set
			{
				if ((this._Ad != value))
				{
					this.OnAdChanging(value);
					this.SendPropertyChanging();
					this._Ad = value;
					this.SendPropertyChanged("Ad");
					this.OnAdChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Firma", DbType="NVarChar(35)")]
		public string Firma
		{
			get
			{
				return this._Firma;
			}
			set
			{
				if ((this._Firma != value))
				{
					this.OnFirmaChanging(value);
					this.SendPropertyChanging();
					this._Firma = value;
					this.SendPropertyChanged("Firma");
					this.OnFirmaChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Tel1", DbType="NVarChar(15)")]
		public string Tel1
		{
			get
			{
				return this._Tel1;
			}
			set
			{
				if ((this._Tel1 != value))
				{
					this.OnTel1Changing(value);
					this.SendPropertyChanging();
					this._Tel1 = value;
					this.SendPropertyChanged("Tel1");
					this.OnTel1Changed();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Tel2", DbType="NVarChar(15)")]
		public string Tel2
		{
			get
			{
				return this._Tel2;
			}
			set
			{
				if ((this._Tel2 != value))
				{
					this.OnTel2Changing(value);
					this.SendPropertyChanging();
					this._Tel2 = value;
					this.SendPropertyChanged("Tel2");
					this.OnTel2Changed();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Adres", DbType="NVarChar(100)")]
		public string Adres
		{
			get
			{
				return this._Adres;
			}
			set
			{
				if ((this._Adres != value))
				{
					this.OnAdresChanging(value);
					this.SendPropertyChanging();
					this._Adres = value;
					this.SendPropertyChanged("Adres");
					this.OnAdresChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Karaliste", DbType="Bit")]
		public System.Nullable<bool> Karaliste
		{
			get
			{
				return this._Karaliste;
			}
			set
			{
				if ((this._Karaliste != value))
				{
					this.OnKaralisteChanging(value);
					this.SendPropertyChanging();
					this._Karaliste = value;
					this.SendPropertyChanged("Karaliste");
					this.OnKaralisteChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.AssociationAttribute(Name="Kisi_Cari", Storage="_Caris", ThisKey="KisiID", OtherKey="KisiID")]
		public EntitySet<Cari> Caris
		{
			get
			{
				return this._Caris;
			}
			set
			{
				this._Caris.Assign(value);
			}
		}
		
		[global::System.Data.Linq.Mapping.AssociationAttribute(Name="Kisi_Senet", Storage="_Senets", ThisKey="KisiID", OtherKey="KisiID")]
		public EntitySet<Senet> Senets
		{
			get
			{
				return this._Senets;
			}
			set
			{
				this._Senets.Assign(value);
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
		
		private void attach_Caris(Cari entity)
		{
			this.SendPropertyChanging();
			entity.Kisi = this;
		}
		
		private void detach_Caris(Cari entity)
		{
			this.SendPropertyChanging();
			entity.Kisi = null;
		}
		
		private void attach_Senets(Senet entity)
		{
			this.SendPropertyChanging();
			entity.Kisi = this;
		}
		
		private void detach_Senets(Senet entity)
		{
			this.SendPropertyChanging();
			entity.Kisi = null;
		}
	}
	
	[global::System.Data.Linq.Mapping.TableAttribute(Name="dbo.Senet")]
	public partial class Senet : INotifyPropertyChanging, INotifyPropertyChanged
	{
		
		private static PropertyChangingEventArgs emptyChangingEventArgs = new PropertyChangingEventArgs(String.Empty);
		
		private int _SenetID;
		
		private int _KisiID;
		
		private string _SeriNo;
		
		private System.Nullable<System.DateTime> _VadeTarihi;
		
		private string _Banka;
		
		private EntityRef<Kisi> _Kisi;
		
    #region Extensibility Method Definitions
    partial void OnLoaded();
    partial void OnValidate(System.Data.Linq.ChangeAction action);
    partial void OnCreated();
    partial void OnSenetIDChanging(int value);
    partial void OnSenetIDChanged();
    partial void OnKisiIDChanging(int value);
    partial void OnKisiIDChanged();
    partial void OnSeriNoChanging(string value);
    partial void OnSeriNoChanged();
    partial void OnVadeTarihiChanging(System.Nullable<System.DateTime> value);
    partial void OnVadeTarihiChanged();
    partial void OnBankaChanging(string value);
    partial void OnBankaChanged();
    #endregion
		
		public Senet()
		{
			this._Kisi = default(EntityRef<Kisi>);
			OnCreated();
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_SenetID", AutoSync=AutoSync.OnInsert, DbType="Int NOT NULL IDENTITY", IsPrimaryKey=true, IsDbGenerated=true)]
		public int SenetID
		{
			get
			{
				return this._SenetID;
			}
			set
			{
				if ((this._SenetID != value))
				{
					this.OnSenetIDChanging(value);
					this.SendPropertyChanging();
					this._SenetID = value;
					this.SendPropertyChanged("SenetID");
					this.OnSenetIDChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_KisiID", DbType="Int NOT NULL")]
		public int KisiID
		{
			get
			{
				return this._KisiID;
			}
			set
			{
				if ((this._KisiID != value))
				{
					if (this._Kisi.HasLoadedOrAssignedValue)
					{
						throw new System.Data.Linq.ForeignKeyReferenceAlreadyHasValueException();
					}
					this.OnKisiIDChanging(value);
					this.SendPropertyChanging();
					this._KisiID = value;
					this.SendPropertyChanged("KisiID");
					this.OnKisiIDChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_SeriNo", DbType="NVarChar(50)")]
		public string SeriNo
		{
			get
			{
				return this._SeriNo;
			}
			set
			{
				if ((this._SeriNo != value))
				{
					this.OnSeriNoChanging(value);
					this.SendPropertyChanging();
					this._SeriNo = value;
					this.SendPropertyChanged("SeriNo");
					this.OnSeriNoChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_VadeTarihi", DbType="Date")]
		public System.Nullable<System.DateTime> VadeTarihi
		{
			get
			{
				return this._VadeTarihi;
			}
			set
			{
				if ((this._VadeTarihi != value))
				{
					this.OnVadeTarihiChanging(value);
					this.SendPropertyChanging();
					this._VadeTarihi = value;
					this.SendPropertyChanged("VadeTarihi");
					this.OnVadeTarihiChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Banka", DbType="NVarChar(50)")]
		public string Banka
		{
			get
			{
				return this._Banka;
			}
			set
			{
				if ((this._Banka != value))
				{
					this.OnBankaChanging(value);
					this.SendPropertyChanging();
					this._Banka = value;
					this.SendPropertyChanged("Banka");
					this.OnBankaChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.AssociationAttribute(Name="Kisi_Senet", Storage="_Kisi", ThisKey="KisiID", OtherKey="KisiID", IsForeignKey=true)]
		public Kisi Kisi
		{
			get
			{
				return this._Kisi.Entity;
			}
			set
			{
				Kisi previousValue = this._Kisi.Entity;
				if (((previousValue != value) 
							|| (this._Kisi.HasLoadedOrAssignedValue == false)))
				{
					this.SendPropertyChanging();
					if ((previousValue != null))
					{
						this._Kisi.Entity = null;
						previousValue.Senets.Remove(this);
					}
					this._Kisi.Entity = value;
					if ((value != null))
					{
						value.Senets.Add(this);
						this._KisiID = value.KisiID;
					}
					else
					{
						this._KisiID = default(int);
					}
					this.SendPropertyChanged("Kisi");
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