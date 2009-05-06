// ===============================================================================
// Microsoft Data Access Application Block for .NET
// http://msdn.microsoft.com/library/en-us/dnbda/html/daab-rm.asp
//
// OleDbHelper.cs
//
// This file contains the implementations of the OleDbHelper and OleDbHelperParameterCache
// classes.
//
// For more information see the Data Access Application Block Implementation Overview. 
// ===============================================================================
// Release history
// VERSION	DESCRIPTION
//   2.0	Added support for FillDataset, UpdateDataset and "Param" helper methods
//
// ===============================================================================
// Copyright (C) 2000-2001 Microsoft Corporation
// All rights reserved.
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY
// OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
// LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR
// FITNESS FOR A PARTICULAR PURPOSE.
// ==============================================================================

using System;
using System.Data;
using System.Xml;
using System.Data.OleDb;
using System.Data.Odbc;
using System.Collections;
using DataAccess;

namespace DataAccess
{
	/// <summary>
	/// The OleDbHelper class is intended to encapsulate high performance, scalable best practices for 
	/// common uses of OleDbClient
	/// </summary>
	public sealed class OleDbHelper
	{
		/// <summary>
		/// Calls the OleDbCommandBuilder.DeriveParameters, doing any setup and cleanup necessary
		/// </summary>
		/// <param name="cmd">The OleDbCommand referencing the stored procedure from which the parameter information is to be derived. The derived parameters are added to the Parameters collection of the OleDbCommand. </param>
		public static void DeriveParameters( OleDbCommand cmd )
		{
			new OleDb().DeriveParameters(cmd);
		}

		#region Private constructor

		// Since this class provides only static methods, make the default constructor private to prevent 
		// instances from being created with "new OleDbHelper()"
		private OleDbHelper() {}

		#endregion Private constructor

		#region GetParameter
		/// <summary>
		/// Get a OleDbParameter for use in a OleDb command
		/// </summary>
		/// <param name="name">The name of the parameter to create</param>
		/// <param name="value">The value of the specified parameter</param>
		/// <returns>A OleDbParameter object</returns>
		public static OleDbParameter GetParameter( string name, object value )
		{
			return (OleDbParameter)(new OleDb().GetParameter(name, value));
		}

		/// <summary>
		/// Get a OleDbParameter for use in a OleDb command
		/// </summary>
		/// <param name="name">The name of the parameter to create</param>
		/// <param name="dbType">The System.Data.DbType of the parameter</param>
		/// <param name="size">The size of the parameter</param>
		/// <param name="direction">The System.Data.ParameterDirection of the parameter</param>
		/// <returns>A OleDbParameter object</returns>
		public static OleDbParameter GetParameter ( string name, DbType dbType, int size, ParameterDirection direction )
		{
			return (OleDbParameter)(new OleDb().GetParameter( name, dbType, size, direction));
		}

		/// <summary>
		/// Get a OleDbParameter for use in a OleDb command
		/// </summary>
		/// <param name="name">The name of the parameter to create</param>
		/// <param name="dbType">The System.Data.DbType of the parameter</param>
		/// <param name="size">The size of the parameter</param>
		/// <param name="sourceColumn">The source column of the parameter</param>
		/// <param name="sourceVersion">The System.Data.DataRowVersion of the parameter</param>
		/// <returns>A OleDbParameter object</returns>
		public static OleDbParameter GetParameter (string name, DbType dbType, int size, string sourceColumn, DataRowVersion sourceVersion )
		{
			return (OleDbParameter)new OleDb().GetParameter(name, dbType, size, sourceColumn, sourceVersion);
		}
		#endregion
		#region ExecuteNonQuery
		/// <summary>
		/// Execute a OleDbCommand (that returns no resultset) against the database specified in 
		/// the connection string
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  int result = ExecuteNonQuery(command);
		/// </remarks>
		/// <param name="command">The OleDbCommand to execute</param>
		/// <returns>An int representing the number of rows affected by the command</returns>
		public static int ExecuteNonQuery(OleDbCommand command)
		{
			return new OleDb().ExecuteNonQuery(command );
		}
		/// <summary>
		/// Execute a OleDbCommand (that returns no resultset and takes no parameters) against the database specified in 
		/// the connection string
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  int result = ExecuteNonQuery(connString, CommandType.StoredProcedure, "PublishOrders");
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <returns>An int representing the number of rows affected by the command</returns>
		public static int ExecuteNonQuery(string connectionString, CommandType commandType, string commandText)
		{
			return new OleDb().ExecuteNonQuery(connectionString, commandType, commandText );
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns no resultset) against the database specified in the connection string 
		/// using the provided parameters
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  int result = ExecuteNonQuery(connString, CommandType.StoredProcedure, "PublishOrders", new OleDbParameter("@prodid", 24));
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="commandParameters">An array of OleDbParamters used to execute the command</param>
		/// <returns>An int representing the number of rows affected by the command</returns>
		public static int ExecuteNonQuery(string connectionString, CommandType commandType, string commandText, params OleDbParameter[] commandParameters)
		{
			return new OleDb().ExecuteNonQuery(connectionString, commandType, commandText, commandParameters);
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns no resultset) against the database specified in 
		/// the connection string using the provided parameter values.  This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <remarks>
		/// This method provides no access to output parameters or the stored procedure's return value parameter.
		/// 
		/// e.g.:  
		///  int result = ExecuteNonQuery(connString, "PublishOrders", 24, 36);
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="spName">The name of the stored prcedure</param>
		/// <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
		/// <returns>An int representing the number of rows affected by the command</returns>
		public static int ExecuteNonQuery(string connectionString, string spName, params object[] parameterValues)
		{
			return new OleDb().ExecuteNonQuery(connectionString, spName, parameterValues );
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns no resultset and takes no parameters) against the provided OleDbConnection. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  int result = ExecuteNonQuery(conn, CommandType.StoredProcedure, "PublishOrders");
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <returns>An int representing the number of rows affected by the command</returns>
		public static int ExecuteNonQuery(OleDbConnection connection, CommandType commandType, string commandText)
		{
			return new OleDb().ExecuteNonQuery(connection, commandType, commandText );
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns no resultset) against the specified OleDbConnection 
		/// using the provided parameters.
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  int result = ExecuteNonQuery(conn, CommandType.StoredProcedure, "PublishOrders", new OleDbParameter("@prodid", 24));
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="commandParameters">An array of OleDbParamters used to execute the command</param>
		/// <returns>An int representing the number of rows affected by the command</returns>
		public static int ExecuteNonQuery(OleDbConnection connection, CommandType commandType, string commandText, params OleDbParameter[] commandParameters)
		{	
			return new OleDb().ExecuteNonQuery(connection, commandType, commandText, commandParameters );
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns no resultset) against the specified OleDbConnection 
		/// using the provided parameter values.  This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <remarks>
		/// This method provides no access to output parameters or the stored procedure's return value parameter.
		/// 
		/// e.g.:  
		///  int result = ExecuteNonQuery(conn, "PublishOrders", 24, 36);
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
		/// <returns>An int representing the number of rows affected by the command</returns>
		public static int ExecuteNonQuery(OleDbConnection connection, string spName, params object[] parameterValues)
		{
			return new OleDb().ExecuteNonQuery(connection, spName, parameterValues );
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns no resultset and takes no parameters) against the provided OleDbTransaction. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  int result = ExecuteNonQuery(trans, CommandType.StoredProcedure, "PublishOrders");
		/// </remarks>
		/// <param name="transaction">A valid OleDbTransaction</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <returns>An int representing the number of rows affected by the command</returns>
		public static int ExecuteNonQuery(OleDbTransaction transaction, CommandType commandType, string commandText)
		{
			return new OleDb().ExecuteNonQuery(transaction, commandType, commandText );
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns no resultset) against the specified OleDbTransaction
		/// using the provided parameters.
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  int result = ExecuteNonQuery(trans, CommandType.StoredProcedure, "GetOrders", new OleDbParameter("@prodid", 24));
		/// </remarks>
		/// <param name="transaction">A valid OleDbTransaction</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="commandParameters">An array of OleDbParamters used to execute the command</param>
		/// <returns>An int representing the number of rows affected by the command</returns>
		public static int ExecuteNonQuery(OleDbTransaction transaction, CommandType commandType, string commandText, params OleDbParameter[] commandParameters)
		{
			return new OleDb().ExecuteNonQuery(transaction, commandType, commandText, commandParameters );
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns no resultset) against the specified 
		/// OleDbTransaction using the provided parameter values.  This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <remarks>
		/// This method provides no access to output parameters or the stored procedure's return value parameter.
		/// 
		/// e.g.:  
		///  int result = ExecuteNonQuery(conn, trans, "PublishOrders", 24, 36);
		/// </remarks>
		/// <param name="transaction">A valid OleDbTransaction</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
		/// <returns>An int representing the number of rows affected by the command</returns>
		public static int ExecuteNonQuery(OleDbTransaction transaction, string spName, params object[] parameterValues)
		{
			return new OleDb().ExecuteNonQuery(transaction, spName, parameterValues );
		}

		#endregion ExecuteNonQuery

		#region ExecuteDataset
		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset) against the database specified in 
		/// the connection string. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  DataSet ds = ExecuteDataset(command);
		/// </remarks>
		/// <param name="command">The OleDbCommand to execute</param>
		/// <returns>A dataset containing the resultset generated by the command</returns>
		public static DataSet ExecuteDataset(OleDbCommand command)
		{
			return new OleDb().ExecuteDataset( command );
		}
		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset and takes no parameters) against the database specified in 
		/// the connection string. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  DataSet ds = ExecuteDataset(connString, CommandType.StoredProcedure, "GetOrders");
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <returns>A dataset containing the resultset generated by the command</returns>
		public static DataSet ExecuteDataset(string connectionString, CommandType commandType, string commandText)
		{
			return new OleDb().ExecuteDataset( connectionString, commandType, commandText );
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset) against the database specified in the connection string 
		/// using the provided parameters.
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  DataSet ds = ExecuteDataset(connString, CommandType.StoredProcedure, "GetOrders", new OleDbParameter("@prodid", 24));
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="commandParameters">An array of OleDbParamters used to execute the command</param>
		/// <returns>A dataset containing the resultset generated by the command</returns>
		public static DataSet ExecuteDataset(string connectionString, CommandType commandType, string commandText, params OleDbParameter[] commandParameters)
		{
			return new OleDb().ExecuteDataset( connectionString, commandType, commandText, commandParameters );
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the database specified in 
		/// the connection string using the provided parameter values.  This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <remarks>
		/// This method provides no access to output parameters or the stored procedure's return value parameter.
		/// 
		/// e.g.:  
		///  DataSet ds = ExecuteDataset(connString, "GetOrders", 24, 36);
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
		/// <returns>A dataset containing the resultset generated by the command</returns>
		public static DataSet ExecuteDataset(string connectionString, string spName, params object[] parameterValues)
		{
			return new OleDb().ExecuteDataset( connectionString, spName, parameterValues );
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset and takes no parameters) against the provided OleDbConnection. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  DataSet ds = ExecuteDataset(conn, CommandType.StoredProcedure, "GetOrders");
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <returns>A dataset containing the resultset generated by the command</returns>
		public static DataSet ExecuteDataset(OleDbConnection connection, CommandType commandType, string commandText)
		{
			return new OleDb().ExecuteDataset( connection, commandType, commandText );
		}
		
		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset) against the specified OleDbConnection 
		/// using the provided parameters.
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  DataSet ds = ExecuteDataset(conn, CommandType.StoredProcedure, "GetOrders", new OleDbParameter("@prodid", 24));
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="commandParameters">An array of OleDbParamters used to execute the command</param>
		/// <returns>A dataset containing the resultset generated by the command</returns>
		public static DataSet ExecuteDataset(OleDbConnection connection, CommandType commandType, string commandText, params OleDbParameter[] commandParameters)
		{
			return new OleDb().ExecuteDataset( connection, commandType, commandText, commandParameters );
		}
		
		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the specified OleDbConnection 
		/// using the provided parameter values.  This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <remarks>
		/// This method provides no access to output parameters or the stored procedure's return value parameter.
		/// 
		/// e.g.:  
		///  DataSet ds = ExecuteDataset(conn, "GetOrders", 24, 36);
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
		/// <returns>A dataset containing the resultset generated by the command</returns>
		public static DataSet ExecuteDataset(OleDbConnection connection, string spName, params object[] parameterValues)
		{
			return new OleDb().ExecuteDataset( connection, spName, parameterValues );
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset and takes no parameters) against the provided OleDbTransaction. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  DataSet ds = ExecuteDataset(trans, CommandType.StoredProcedure, "GetOrders");
		/// </remarks>
		/// <param name="transaction">A valid OleDbTransaction</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <returns>A dataset containing the resultset generated by the command</returns>
		public static DataSet ExecuteDataset(OleDbTransaction transaction, CommandType commandType, string commandText)
		{
			return new OleDb().ExecuteDataset( transaction, commandType, commandText );
		}
		
		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset) against the specified OleDbTransaction
		/// using the provided parameters.
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  DataSet ds = ExecuteDataset(trans, CommandType.StoredProcedure, "GetOrders", new OleDbParameter("@prodid", 24));
		/// </remarks>
		/// <param name="transaction">A valid OleDbTransaction</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="commandParameters">An array of OleDbParamters used to execute the command</param>
		/// <returns>A dataset containing the resultset generated by the command</returns>
		public static DataSet ExecuteDataset(OleDbTransaction transaction, CommandType commandType, string commandText, params OleDbParameter[] commandParameters)
		{
			return new OleDb().ExecuteDataset( transaction, commandType, commandText, commandParameters );
		}
		
		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the specified 
		/// OleDbTransaction using the provided parameter values.  This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <remarks>
		/// This method provides no access to output parameters or the stored procedure's return value parameter.
		/// 
		/// e.g.:  
		///  DataSet ds = ExecuteDataset(trans, "GetOrders", 24, 36);
		/// </remarks>
		/// <param name="transaction">A valid OleDbTransaction</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
		/// <returns>A dataset containing the resultset generated by the command</returns>
		public static DataSet ExecuteDataset(OleDbTransaction transaction, string spName, params object[] parameterValues)
		{
			return new OleDb().ExecuteDataset( transaction, spName, parameterValues );
		}

		#endregion ExecuteDataset
		
		#region ExecuteReader
		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset) against the database specified in 
		/// the connection string. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  OleDbDataReader dr = ExecuteReader(command);
		/// </remarks>
		/// <param name="command">The OleDbCommand to execute</param>
		/// <returns>A OleDbDataReader containing the resultset generated by the command</returns>
		public static OleDbDataReader ExecuteReader(OleDbCommand command)
		{
			return new OleDb().ExecuteReader( command) as OleDbDataReader;
		}
		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset and takes no parameters) against the database specified in 
		/// the connection string. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  OleDbDataReader dr = ExecuteReader(connString, CommandType.StoredProcedure, "GetOrders");
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <returns>A OleDbDataReader containing the resultset generated by the command</returns>
		public static OleDbDataReader ExecuteReader(string connectionString, CommandType commandType, string commandText)
		{
			return new OleDb().ExecuteReader( connectionString, commandType, commandText) as OleDbDataReader;
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset) against the database specified in the connection string 
		/// using the provided parameters.
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  OleDbDataReader dr = ExecuteReader(connString, CommandType.StoredProcedure, "GetOrders", new OleDbParameter("@prodid", 24));
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="commandParameters">An array of OleDbParamters used to execute the command</param>
		/// <returns>A OleDbDataReader containing the resultset generated by the command</returns>
		public static OleDbDataReader ExecuteReader(string connectionString, CommandType commandType, string commandText, params OleDbParameter[] commandParameters)
		{
			return new OleDb().ExecuteReader( connectionString, commandType, commandText, commandParameters ) as OleDbDataReader;
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the database specified in 
		/// the connection string using the provided parameter values.  This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <remarks>
		/// This method provides no access to output parameters or the stored procedure's return value parameter.
		/// 
		/// e.g.:  
		///  OleDbDataReader dr = ExecuteReader(connString, "GetOrders", 24, 36);
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
		/// <returns>A OleDbDataReader containing the resultset generated by the command</returns>
		public static OleDbDataReader ExecuteReader(string connectionString, string spName, params object[] parameterValues)
		{
			return new OleDb().ExecuteReader( connectionString, spName, parameterValues ) as OleDbDataReader;
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset and takes no parameters) against the provided OleDbConnection. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  OleDbDataReader dr = ExecuteReader(conn, CommandType.StoredProcedure, "GetOrders");
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <returns>A OleDbDataReader containing the resultset generated by the command</returns>
		public static OleDbDataReader ExecuteReader(OleDbConnection connection, CommandType commandType, string commandText)
		{
			return new OleDb().ExecuteReader( connection, commandType, commandText ) as OleDbDataReader;
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset) against the specified OleDbConnection 
		/// using the provided parameters.
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  OleDbDataReader dr = ExecuteReader(conn, CommandType.StoredProcedure, "GetOrders", new OleDbParameter("@prodid", 24));
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="commandParameters">An array of OleDbParamters used to execute the command</param>
		/// <returns>A OleDbDataReader containing the resultset generated by the command</returns>
		public static OleDbDataReader ExecuteReader(OleDbConnection connection, CommandType commandType, string commandText, params OleDbParameter[] commandParameters)
		{
			return new OleDb().ExecuteReader( connection, commandType, commandText, commandParameters) as OleDbDataReader;
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the specified OleDbConnection 
		/// using the provided parameter values.  This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <remarks>
		/// This method provides no access to output parameters or the stored procedure's return value parameter.
		/// 
		/// e.g.:  
		///  OleDbDataReader dr = ExecuteReader(conn, "GetOrders", 24, 36);
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
		/// <returns>A OleDbDataReader containing the resultset generated by the command</returns>
		public static OleDbDataReader ExecuteReader(OleDbConnection connection, string spName, params object[] parameterValues)
		{
			return new OleDb().ExecuteReader( connection, spName, parameterValues ) as OleDbDataReader;
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset and takes no parameters) against the provided OleDbTransaction. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  OleDbDataReader dr = ExecuteReader(trans, CommandType.StoredProcedure, "GetOrders");
		/// </remarks>
		/// <param name="transaction">A valid OleDbTransaction</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <returns>A OleDbDataReader containing the resultset generated by the command</returns>
		public static OleDbDataReader ExecuteReader(OleDbTransaction transaction, CommandType commandType, string commandText)
		{
			return new OleDb().ExecuteReader( transaction, commandType, commandText ) as OleDbDataReader;
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset) against the specified OleDbTransaction
		/// using the provided parameters.
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///   OleDbDataReader dr = ExecuteReader(trans, CommandType.StoredProcedure, "GetOrders", new OleDbParameter("@prodid", 24));
		/// </remarks>
		/// <param name="transaction">A valid OleDbTransaction</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="commandParameters">An array of OleDbParamters used to execute the command</param>
		/// <returns>A OleDbDataReader containing the resultset generated by the command</returns>
		public static OleDbDataReader ExecuteReader(OleDbTransaction transaction, CommandType commandType, string commandText, params OleDbParameter[] commandParameters)
		{
			return new OleDb().ExecuteReader( transaction, commandType, commandText, commandParameters ) as OleDbDataReader;
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the specified
		/// OleDbTransaction using the provided parameter values.  This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <remarks>
		/// This method provides no access to output parameters or the stored procedure's return value parameter.
		/// 
		/// e.g.:  
		///  OleDbDataReader dr = ExecuteReader(trans, "GetOrders", 24, 36);
		/// </remarks>
		/// <param name="transaction">A valid OleDbTransaction</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
		/// <returns>A OleDbDataReader containing the resultset generated by the command</returns>
		public static OleDbDataReader ExecuteReader(OleDbTransaction transaction, string spName, params object[] parameterValues)
		{
			return new OleDb().ExecuteReader( transaction, spName, parameterValues ) as OleDbDataReader;
		}

		#endregion ExecuteReader

		#region ExecuteScalar
		/// <summary>
		/// Execute a OleDbCommand (that returns a 1x1 resultset) against the database specified in 
		/// the connection string. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  int orderCount = (int)ExecuteScalar(command);
		/// </remarks>
		/// <param name="command">The OleDbCommand to execute</param>
		/// <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
		public static object ExecuteScalar(OleDbCommand command)
		{
			// Pass through the call providing null for the set of OleDbParameters
			return new OleDb().ExecuteScalar( command);
		}
		/// <summary>
		/// Execute a OleDbCommand (that returns a 1x1 resultset and takes no parameters) against the database specified in 
		/// the connection string. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  int orderCount = (int)ExecuteScalar(connString, CommandType.StoredProcedure, "GetOrderCount");
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
		public static object ExecuteScalar(string connectionString, CommandType commandType, string commandText)
		{
			// Pass through the call providing null for the set of OleDbParameters
			return new OleDb().ExecuteScalar( connectionString, commandType, commandText);
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a 1x1 resultset) against the database specified in the connection string 
		/// using the provided parameters.
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  int orderCount = (int)ExecuteScalar(connString, CommandType.StoredProcedure, "GetOrderCount", new OleDbParameter("@prodid", 24));
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="commandParameters">An array of OleDbParamters used to execute the command</param>
		/// <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
		public static object ExecuteScalar(string connectionString, CommandType commandType, string commandText, params OleDbParameter[] commandParameters)
		{
			return new OleDb().ExecuteScalar( connectionString, commandType, commandText, commandParameters);
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a 1x1 resultset) against the database specified in 
		/// the connection string using the provided parameter values.  This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <remarks>
		/// This method provides no access to output parameters or the stored procedure's return value parameter.
		/// 
		/// e.g.:  
		///  int orderCount = (int)ExecuteScalar(connString, "GetOrderCount", 24, 36);
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
		/// <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
		public static object ExecuteScalar(string connectionString, string spName, params object[] parameterValues)
		{
			return new OleDb().ExecuteScalar( connectionString, spName, parameterValues );
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a 1x1 resultset and takes no parameters) against the provided OleDbConnection. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  int orderCount = (int)ExecuteScalar(conn, CommandType.StoredProcedure, "GetOrderCount");
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
		public static object ExecuteScalar(OleDbConnection connection, CommandType commandType, string commandText)
		{
			return new OleDb().ExecuteScalar( connection, commandType, commandText );
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a 1x1 resultset) against the specified OleDbConnection 
		/// using the provided parameters.
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  int orderCount = (int)ExecuteScalar(conn, CommandType.StoredProcedure, "GetOrderCount", new OleDbParameter("@prodid", 24));
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="commandParameters">An array of OleDbParamters used to execute the command</param>
		/// <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
		public static object ExecuteScalar(OleDbConnection connection, CommandType commandType, string commandText, params OleDbParameter[] commandParameters)
		{
			return new OleDb().ExecuteScalar( connection, commandType, commandText, commandParameters );
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a 1x1 resultset) against the specified OleDbConnection 
		/// using the provided parameter values.  This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <remarks>
		/// This method provides no access to output parameters or the stored procedure's return value parameter.
		/// 
		/// e.g.:  
		///  int orderCount = (int)ExecuteScalar(conn, "GetOrderCount", 24, 36);
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
		/// <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
		public static object ExecuteScalar(OleDbConnection connection, string spName, params object[] parameterValues)
		{
			return new OleDb().ExecuteScalar( connection, spName, parameterValues );
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a 1x1 resultset and takes no parameters) against the provided OleDbTransaction. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  int orderCount = (int)ExecuteScalar(trans, CommandType.StoredProcedure, "GetOrderCount");
		/// </remarks>
		/// <param name="transaction">A valid OleDbTransaction</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
		public static object ExecuteScalar(OleDbTransaction transaction, CommandType commandType, string commandText)
		{
			return new OleDb().ExecuteScalar( transaction, commandType, commandText );
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a 1x1 resultset) against the specified OleDbTransaction
		/// using the provided parameters.
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  int orderCount = (int)ExecuteScalar(trans, CommandType.StoredProcedure, "GetOrderCount", new OleDbParameter("@prodid", 24));
		/// </remarks>
		/// <param name="transaction">A valid OleDbTransaction</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="commandParameters">An array of OleDbParamters used to execute the command</param>
		/// <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
		public static object ExecuteScalar(OleDbTransaction transaction, CommandType commandType, string commandText, params OleDbParameter[] commandParameters)
		{
			return new OleDb().ExecuteScalar( transaction, commandType, commandText, commandParameters );
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a 1x1 resultset) against the specified
		/// OleDbTransaction using the provided parameter values.  This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <remarks>
		/// This method provides no access to output parameters or the stored procedure's return value parameter.
		/// 
		/// e.g.:  
		///  int orderCount = (int)ExecuteScalar(trans, "GetOrderCount", 24, 36);
		/// </remarks>
		/// <param name="transaction">A valid OleDbTransaction</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
		/// <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
		public static object ExecuteScalar(OleDbTransaction transaction, string spName, params object[] parameterValues)
		{
			return new OleDb().ExecuteScalar( transaction, spName, parameterValues );
		}

		#endregion ExecuteScalar	

		#region ExecuteXmlReader
		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset) against the provided OleDbConnection. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  XmlReader r = ExecuteXmlReader(OleDbCommand command);
		/// </remarks>
		/// <param name="command">The OleDbCommand to execute</param>
		/// <returns>An XmlReader containing the resultset generated by the command</returns>
		public static XmlReader ExecuteXmlReader(OleDbCommand command)
		{
			return new OleDb().ExecuteXmlReader( command);
		}
		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset and takes no parameters) against the provided OleDbConnection. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  XmlReader r = ExecuteXmlReader(conn, CommandType.StoredProcedure, "GetOrders");
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command using "FOR XML AUTO"</param>
		/// <returns>An XmlReader containing the resultset generated by the command</returns>
		public static XmlReader ExecuteXmlReader(OleDbConnection connection, CommandType commandType, string commandText)
		{
			return new OleDb().ExecuteXmlReader( connection, commandType, commandText);
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset) against the specified OleDbConnection 
		/// using the provided parameters.
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  XmlReader r = ExecuteXmlReader(conn, CommandType.StoredProcedure, "GetOrders", new OleDbParameter("@prodid", 24));
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command using "FOR XML AUTO"</param>
		/// <param name="commandParameters">An array of OleDbParamters used to execute the command</param>
		/// <returns>An XmlReader containing the resultset generated by the command</returns>
		public static XmlReader ExecuteXmlReader(OleDbConnection connection, CommandType commandType, string commandText, params OleDbParameter[] commandParameters)
		{
			return new OleDb().ExecuteXmlReader( connection, commandType, commandText, commandParameters );
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the specified OleDbConnection 
		/// using the provided parameter values.  This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <remarks>
		/// This method provides no access to output parameters or the stored procedure's return value parameter.
		/// 
		/// e.g.:  
		///  XmlReader r = ExecuteXmlReader(conn, "GetOrders", 24, 36);
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection</param>
		/// <param name="spName">The name of the stored procedure using "FOR XML AUTO"</param>
		/// <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
		/// <returns>An XmlReader containing the resultset generated by the command</returns>
		public static XmlReader ExecuteXmlReader(OleDbConnection connection, string spName, params object[] parameterValues)
		{
			return new OleDb().ExecuteXmlReader( connection, spName, parameterValues );
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset and takes no parameters) against the provided OleDbTransaction. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  XmlReader r = ExecuteXmlReader(trans, CommandType.StoredProcedure, "GetOrders");
		/// </remarks>
		/// <param name="transaction">A valid OleDbTransaction</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command using "FOR XML AUTO"</param>
		/// <returns>An XmlReader containing the resultset generated by the command</returns>
		public static XmlReader ExecuteXmlReader(OleDbTransaction transaction, CommandType commandType, string commandText)
		{
			return new OleDb().ExecuteXmlReader( transaction, commandType, commandText );
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset) against the specified OleDbTransaction
		/// using the provided parameters.
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  XmlReader r = ExecuteXmlReader(trans, CommandType.StoredProcedure, "GetOrders", new OleDbParameter("@prodid", 24));
		/// </remarks>
		/// <param name="transaction">A valid OleDbTransaction</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command using "FOR XML AUTO"</param>
		/// <param name="commandParameters">An array of OleDbParamters used to execute the command</param>
		/// <returns>An XmlReader containing the resultset generated by the command</returns>
		public static XmlReader ExecuteXmlReader(OleDbTransaction transaction, CommandType commandType, string commandText, params OleDbParameter[] commandParameters)
		{
			return new OleDb().ExecuteXmlReader( transaction, commandType, commandText, commandParameters );
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the specified 
		/// OleDbTransaction using the provided parameter values.  This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <remarks>
		/// This method provides no access to output parameters or the stored procedure's return value parameter.
		/// 
		/// e.g.:  
		///  XmlReader r = ExecuteXmlReader(trans, "GetOrders", 24, 36);
		/// </remarks>
		/// <param name="transaction">A valid OleDbTransaction</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
		/// <returns>A dataset containing the resultset generated by the command</returns>
		public static XmlReader ExecuteXmlReader(OleDbTransaction transaction, string spName, params object[] parameterValues)
		{
			return new OleDb().ExecuteXmlReader( transaction, spName, parameterValues );
		}

		#endregion ExecuteXmlReader

		#region FillDataset
		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset) against the database specified in 
		/// the connection string. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  FillDataset(connString, CommandType.StoredProcedure, "GetOrders", ds, new string[] {"orders"});
		/// </remarks>
		/// <param name="command">The OleDbCommand to execute</param>
		/// <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
		/// <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced
		/// by a user defined name (probably the actual table name)</param>
		public static void FillDataset(OleDbCommand command, DataSet dataSet, string[] tableNames)
		{
			new OleDb().FillDataset( command, dataSet, tableNames);
		}
		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset and takes no parameters) against the database specified in 
		/// the connection string. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  FillDataset(connString, CommandType.StoredProcedure, "GetOrders", ds, new string[] {"orders"});
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
		/// <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced
		/// by a user defined name (probably the actual table name)</param>
		public static void FillDataset(string connectionString, CommandType commandType, string commandText, DataSet dataSet, string[] tableNames)
		{
			new OleDb().FillDataset( connectionString, commandType, commandText, dataSet, tableNames);
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset) against the database specified in the connection string 
		/// using the provided parameters.
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  FillDataset(connString, CommandType.StoredProcedure, "GetOrders", ds, new string[] {"orders"}, new OleDbParameter("@prodid", 24));
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="commandParameters">An array of OleDbParamters used to execute the command</param>
		/// <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
		/// <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced
		/// by a user defined name (probably the actual table name)
		/// </param>
		public static void FillDataset(string connectionString, CommandType commandType,
			string commandText, DataSet dataSet, string[] tableNames,
			params OleDbParameter[] commandParameters)
		{
			new OleDb().FillDataset( connectionString, commandType, commandText, dataSet, tableNames, commandParameters );
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the database specified in 
		/// the connection string using the provided parameter values.  This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <remarks>
		/// This method provides no access to output parameters or the stored procedure's return value parameter.
		/// 
		/// e.g.:  
		///  FillDataset(connString, CommandType.StoredProcedure, "GetOrders", ds, new string[] {"orders"}, 24);
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
		/// <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced
		/// by a user defined name (probably the actual table name)
		/// </param>    
		/// <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
		public static void FillDataset(string connectionString, string spName,
			DataSet dataSet, string[] tableNames, params object[] parameterValues)
		{
			new OleDb().FillDataset( connectionString, spName, dataSet, tableNames, parameterValues );
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset and takes no parameters) against the provided OleDbConnection. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  FillDataset(conn, CommandType.StoredProcedure, "GetOrders", ds, new string[] {"orders"});
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
		/// <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced
		/// by a user defined name (probably the actual table name)
		/// </param>    
		public static void FillDataset(OleDbConnection connection, CommandType commandType, 
			string commandText, DataSet dataSet, string[] tableNames)
		{
			new OleDb().FillDataset( connection, commandType, commandText, dataSet, tableNames );
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset) against the specified OleDbConnection 
		/// using the provided parameters.
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  FillDataset(conn, CommandType.StoredProcedure, "GetOrders", ds, new string[] {"orders"}, new OleDbParameter("@prodid", 24));
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
		/// <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced
		/// by a user defined name (probably the actual table name)
		/// </param>
		/// <param name="commandParameters">An array of OleDbParamters used to execute the command</param>
		public static void FillDataset(OleDbConnection connection, CommandType commandType, 
			string commandText, DataSet dataSet, string[] tableNames, params OleDbParameter[] commandParameters)
		{
			new OleDb().FillDataset( connection, commandType, commandText, dataSet, tableNames, commandParameters );
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the specified OleDbConnection 
		/// using the provided parameter values.  This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <remarks>
		/// This method provides no access to output parameters or the stored procedure's return value parameter.
		/// 
		/// e.g.:  
		///  FillDataset(conn, "GetOrders", ds, new string[] {"orders"}, 24, 36);
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
		/// <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced
		/// by a user defined name (probably the actual table name)
		/// </param>
		/// <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
		public static void FillDataset(OleDbConnection connection, string spName, 
			DataSet dataSet, string[] tableNames, params object[] parameterValues)
		{
			new OleDb().FillDataset( connection, spName, dataSet, tableNames, parameterValues );
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset and takes no parameters) against the provided OleDbTransaction. 
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  FillDataset(trans, CommandType.StoredProcedure, "GetOrders", ds, new string[] {"orders"});
		/// </remarks>
		/// <param name="transaction">A valid OleDbTransaction</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
		/// <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced
		/// by a user defined name (probably the actual table name)
		/// </param>
		public static void FillDataset(OleDbTransaction transaction, CommandType commandType, 
			string commandText, DataSet dataSet, string[] tableNames)
		{
			new OleDb().FillDataset( transaction, commandType, commandText, dataSet, tableNames );
		}

		/// <summary>
		/// Execute a OleDbCommand (that returns a resultset) against the specified OleDbTransaction
		/// using the provided parameters.
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  FillDataset(trans, CommandType.StoredProcedure, "GetOrders", ds, new string[] {"orders"}, new OleDbParameter("@prodid", 24));
		/// </remarks>
		/// <param name="transaction">A valid OleDbTransaction</param>
		/// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
		/// <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced
		/// by a user defined name (probably the actual table name)
		/// </param>
		/// <param name="commandParameters">An array of OleDbParamters used to execute the command</param>
		public static void FillDataset(OleDbTransaction transaction, CommandType commandType, 
			string commandText, DataSet dataSet, string[] tableNames, params OleDbParameter[] commandParameters)
		{
			new OleDb().FillDataset( transaction, commandType, commandText, dataSet, tableNames, commandParameters );
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the specified 
		/// OleDbTransaction using the provided parameter values.  This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <remarks>
		/// This method provides no access to output parameters or the stored procedure's return value parameter.
		/// 
		/// e.g.:  
		///  FillDataset(trans, "GetOrders", ds, new string[]{"orders"}, 24, 36);
		/// </remarks>
		/// <param name="transaction">A valid OleDbTransaction</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
		/// <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced
		/// by a user defined name (probably the actual table name)
		/// </param>
		/// <param name="parameterValues">An array of objects to be assigned as the input values of the stored procedure</param>
		public static void FillDataset(OleDbTransaction transaction, string spName,
			DataSet dataSet, string[] tableNames, params object[] parameterValues) 
		{
			new OleDb().FillDataset( transaction, spName, dataSet, tableNames, parameterValues );
		}

		#endregion
        
		#region UpdateDataset
		/// <summary>
		/// Executes the respective command for each inserted, updated, or deleted row in the DataSet.
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  UpdateDataset(conn, insertCommand, deleteCommand, updateCommand, dataSet, "Order");
		/// </remarks>
		/// <param name="insertCommand">A valid transact-OleDb statement or stored procedure to insert new records into the data source</param>
		/// <param name="deleteCommand">A valid transact-OleDb statement or stored procedure to delete records from the data source</param>
		/// <param name="updateCommand">A valid transact-OleDb statement or stored procedure used to update records in the data source</param>
		/// <param name="dataSet">The DataSet used to update the data source</param>
		/// <param name="tableName">The DataTable used to update the data source.</param>
		public static void UpdateDataset(OleDbCommand insertCommand, OleDbCommand deleteCommand, OleDbCommand updateCommand, DataSet dataSet, string tableName)
		{
			new OleDb().UpdateDataset( insertCommand, deleteCommand, updateCommand, dataSet, tableName);
		}

		/// <summary> 
		/// Executes the System.Data.OleDbClient.OleDbCommand for each inserted, updated, or deleted row in the DataSet also implementing RowUpdating and RowUpdated Event Handlers 
		/// </summary> 
		/// <remarks> 
		/// e.g.:  
		/// OleDbRowUpdatingEventHandler rowUpdating = new OleDbRowUpdatingEventHandler( OnRowUpdating ); 
		/// OleDbRowUpdatedEventHandler rowUpdated = new OleDbRowUpdatedEventHandler( OnRowUpdated ); 
		/// adoHelper.UpdateDataSet(OleDbInsertCommand, OleDbDeleteCommand, OleDbUpdateCommand, dataSet, "Order", rowUpdating, rowUpdated); 
		/// </remarks> 
		/// <param name="insertCommand">A valid transact-OleDb statement or stored procedure to insert new records into the data source</param> 
		/// <param name="deleteCommand">A valid transact-OleDb statement or stored procedure to delete records from the data source</param> 
		/// <param name="updateCommand">A valid transact-OleDb statement or stored procedure used to update records in the data source</param> 
		/// <param name="dataSet">The DataSet used to update the data source</param> 
		/// <param name="tableName">The DataTable used to update the data source.</param> 
		/// <param name="rowUpdating">The AdoHelper.RowUpdatingEventHandler or null</param> 
		/// <param name="rowUpdated">The AdoHelper.RowUpdatedEventHandler or null</param> 
		public static void UpdateDataset(IDbCommand insertCommand, IDbCommand deleteCommand, IDbCommand updateCommand, 
			DataSet dataSet, string tableName, AdoHelper.RowUpdatingHandler rowUpdating, AdoHelper.RowUpdatedHandler rowUpdated) 
		{
			new OleDb().UpdateDataset( insertCommand, deleteCommand, updateCommand, dataSet, tableName, rowUpdating, rowUpdated);
		}
		
		#endregion

		#region CreateCommand
		/// <summary>
		/// Simplify the creation of a OleDb command object by allowing
		/// a stored procedure and optional parameters to be provided
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  OleDbCommand command = CreateCommand(connenctionString, "AddCustomer", "CustomerID", "CustomerName");
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="sourceColumns">An array of string to be assigned as the source columns of the stored procedure parameters</param>
		/// <returns>A valid OleDbCommand object</returns>
		public static OleDbCommand CreateCommand(string connectionString, string spName, params string[] sourceColumns) 
		{
			return new OleDb().CreateCommand( connectionString, spName, sourceColumns ) as OleDbCommand;
		}
		/// <summary>
		/// Simplify the creation of a OleDb command object by allowing
		/// a stored procedure and optional parameters to be provided
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  OleDbCommand command = CreateCommand(conn, "AddCustomer", "CustomerID", "CustomerName");
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection object</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="sourceColumns">An array of string to be assigned as the source columns of the stored procedure parameters</param>
		/// <returns>A valid OleDbCommand object</returns>
		public static OleDbCommand CreateCommand(OleDbConnection connection, string spName, params string[] sourceColumns) 
		{
			return new OleDb().CreateCommand( connection, spName, sourceColumns ) as OleDbCommand;
		}
		/// <summary>
		/// Simplify the creation of a OleDb command object by allowing
		/// a stored procedure and optional parameters to be provided
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  OleDbCommand command = CreateCommand(connenctionString, "AddCustomer", "CustomerID", "CustomerName");
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="commandText">A valid OleDb string to execute</param>
		/// <param name="commandType">The CommandType to execute (i.e. StoredProcedure, Text)</param>
		/// <param name="commandParameters">The OleDbParameters to pass to the command</param>
		/// <returns>A valid OleDbCommand object</returns>
		public static OleDbCommand CreateCommand(string connectionString, string commandText, CommandType commandType, params OleDbParameter[] commandParameters) 
		{
			return new OleDb().CreateCommand( connectionString, commandText, commandType, commandParameters ) as OleDbCommand;
		}
		/// <summary>
		/// Simplify the creation of a OleDb command object by allowing
		/// a stored procedure and optional parameters to be provided
		/// </summary>
		/// <remarks>
		/// e.g.:  
		///  OleDbCommand command = CreateCommand(conn, "AddCustomer", "CustomerID", "CustomerName");
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection object</param>
		/// <param name="commandText">A valid OleDb string to execute</param>
		/// <param name="commandType">The CommandType to execute (i.e. StoredProcedure, Text)</param>
		/// <param name="commandParameters">The OleDbParameters to pass to the command</param>
		/// <returns>A valid OleDbCommand object</returns>
		public static OleDbCommand CreateCommand(OleDbConnection connection, string commandText, CommandType commandType, params OleDbParameter[] commandParameters) 
		{
			return new OleDb().CreateCommand( connection, commandText, commandType, commandParameters) as OleDbCommand;
		}
		#endregion

		#region ExecuteNonQueryTypedParams
		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns no resultset) against the database specified in 
		/// the connection string using the dataRow column values as the stored procedure's parameters values.
		/// This method will assign the parameter values based on row values.
		/// </summary>
		/// <param name="command">The OleDbCommand to execute</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>An int representing the number of rows affected by the command</returns>
		public static int ExecuteNonQueryTypedParams(OleDbCommand command, DataRow dataRow)
		{
			return new OleDb().ExecuteNonQueryTypedParams( command, dataRow);
		}
		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns no resultset) against the database specified in 
		/// the connection string using the dataRow column values as the stored procedure's parameters values.
		/// This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on row values.
		/// </summary>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>An int representing the number of rows affected by the command</returns>
		public static int ExecuteNonQueryTypedParams(String connectionString, String spName, DataRow dataRow)
		{
			return new OleDb().ExecuteNonQueryTypedParams( connectionString, spName, dataRow);
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns no resultset) against the specified OleDbConnection 
		/// using the dataRow column values as the stored procedure's parameters values.  
		/// This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on row values.
		/// </summary>
		/// <param name="connection">A valid OleDbConnection object</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>An int representing the number of rows affected by the command</returns>
		public static int ExecuteNonQueryTypedParams(OleDbConnection connection, String spName, DataRow dataRow)
		{
			return new OleDb().ExecuteNonQueryTypedParams( connection, spName, dataRow);
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns no resultset) against the specified
		/// OleDbTransaction using the dataRow column values as the stored procedure's parameters values.
		/// This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on row values.
		/// </summary>
		/// <param name="transaction">A valid OleDbTransaction object</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>An int representing the number of rows affected by the command</returns>
		public static int ExecuteNonQueryTypedParams(OleDbTransaction transaction, String spName, DataRow dataRow)
		{
			return new OleDb().ExecuteNonQueryTypedParams( transaction, spName, dataRow);
		}
		#endregion

		#region ExecuteDatasetTypedParams
		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the database specified in 
		/// the connection string using the dataRow column values as the stored procedure's parameters values.
		/// This method will assign the parameter values based on row values.
		/// </summary>
		/// <param name="command">The OleDbCommand to execute</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>A dataset containing the resultset generated by the command</returns>
		public static DataSet ExecuteDatasetTypedParams(OleDbCommand command, DataRow dataRow)
		{
			return new OleDb().ExecuteDatasetTypedParams( command, dataRow );
		}
		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the database specified in 
		/// the connection string using the dataRow column values as the stored procedure's parameters values.
		/// This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on row values.
		/// </summary>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>A dataset containing the resultset generated by the command</returns>
		public static DataSet ExecuteDatasetTypedParams(string connectionString, String spName, DataRow dataRow)
		{
			return new OleDb().ExecuteDatasetTypedParams( connectionString, spName, dataRow );
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the specified OleDbConnection 
		/// using the dataRow column values as the store procedure's parameters values.
		/// This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on row values.
		/// </summary>
		/// <param name="connection">A valid OleDbConnection object</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>A dataset containing the resultset generated by the command</returns>
		public static DataSet ExecuteDatasetTypedParams(OleDbConnection connection, String spName, DataRow dataRow)
		{
			return new OleDb().ExecuteDatasetTypedParams( connection, spName, dataRow );
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the specified OleDbTransaction 
		/// using the dataRow column values as the stored procedure's parameters values.
		/// This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on row values.
		/// </summary>
		/// <param name="transaction">A valid OleDbTransaction object</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>A dataset containing the resultset generated by the command</returns>
		public static DataSet ExecuteDatasetTypedParams(OleDbTransaction transaction, String spName, DataRow dataRow)
		{
			return new OleDb().ExecuteDatasetTypedParams( transaction, spName, dataRow );
		}

		#endregion

		#region ExecuteReaderTypedParams
		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the database specified in 
		/// the connection string using the dataRow column values as the stored procedure's parameters values.
		/// This method will assign the parameter values based on parameter order.
		/// </summary>
		/// <param name="command">The OleDbCommand toe execute</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>A OleDbDataReader containing the resultset generated by the command</returns>
		public static OleDbDataReader ExecuteReaderTypedParams(OleDbCommand command, DataRow dataRow)
		{
			return new OleDb().ExecuteReaderTypedParams( command, dataRow ) as OleDbDataReader;
		}
		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the database specified in 
		/// the connection string using the dataRow column values as the stored procedure's parameters values.
		/// This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>A OleDbDataReader containing the resultset generated by the command</returns>
		public static OleDbDataReader ExecuteReaderTypedParams(String connectionString, String spName, DataRow dataRow)
		{
			return new OleDb().ExecuteReaderTypedParams( connectionString, spName, dataRow ) as OleDbDataReader;
		}

                
		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the specified OleDbConnection 
		/// using the dataRow column values as the stored procedure's parameters values.
		/// This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <param name="connection">A valid OleDbConnection object</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>A OleDbDataReader containing the resultset generated by the command</returns>
		public static OleDbDataReader ExecuteReaderTypedParams(OleDbConnection connection, String spName, DataRow dataRow)
		{
			return new OleDb().ExecuteReaderTypedParams( connection, spName, dataRow ) as OleDbDataReader;
		}
        
		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the specified OleDbTransaction 
		/// using the dataRow column values as the stored procedure's parameters values.
		/// This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <param name="transaction">A valid OleDbTransaction object</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>A OleDbDataReader containing the resultset generated by the command</returns>
		public static OleDbDataReader ExecuteReaderTypedParams(OleDbTransaction transaction, String spName, DataRow dataRow)
		{
			return new OleDb().ExecuteReaderTypedParams( transaction, spName, dataRow ) as OleDbDataReader;
		}
		#endregion

		#region ExecuteScalarTypedParams
		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a 1x1 resultset) against the database specified in 
		/// the connection string using the dataRow column values as the stored procedure's parameters values.
		/// This method will assign the parameter values based on parameter order.
		/// </summary>
		/// <param name="command">The OleDbCommand to execute</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
		public static object ExecuteScalarTypedParams(OleDbCommand command, DataRow dataRow)
		{
			return new OleDb().ExecuteScalarTypedParams( command, dataRow );
		}
		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a 1x1 resultset) against the database specified in 
		/// the connection string using the dataRow column values as the stored procedure's parameters values.
		/// This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
		public static object ExecuteScalarTypedParams(String connectionString, String spName, DataRow dataRow)
		{
			return new OleDb().ExecuteScalarTypedParams( connectionString, spName, dataRow );
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a 1x1 resultset) against the specified OleDbConnection 
		/// using the dataRow column values as the stored procedure's parameters values.
		/// This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <param name="connection">A valid OleDbConnection object</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
		public static object ExecuteScalarTypedParams(OleDbConnection connection, String spName, DataRow dataRow)
		{
			return new OleDb().ExecuteScalarTypedParams( connection, spName, dataRow );
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a 1x1 resultset) against the specified OleDbTransaction
		/// using the dataRow column values as the stored procedure's parameters values.
		/// This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <param name="transaction">A valid OleDbTransaction object</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>An object containing the value in the 1x1 resultset generated by the command</returns>
		public static object ExecuteScalarTypedParams(OleDbTransaction transaction, String spName, DataRow dataRow)
		{
			return new OleDb().ExecuteScalarTypedParams( transaction, spName, dataRow );
		}
		#endregion

		#region ExecuteXmlReaderTypedParams
		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the specified OleDbConnection 
		/// using the dataRow column values as the stored procedure's parameters values.
		/// This method will assign the parameter values based on parameter order.
		/// </summary>
		/// <param name="command">The OleDbCommand to execute</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>An XmlReader containing the resultset generated by the command</returns>
		public static XmlReader ExecuteXmlReaderTypedParams(OleDbCommand command, DataRow dataRow)
		{
			return new OleDb().ExecuteXmlReaderTypedParams( command, dataRow );
		}
		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the specified OleDbConnection 
		/// using the dataRow column values as the stored procedure's parameters values.
		/// This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <param name="connection">A valid OleDbConnection object</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>An XmlReader containing the resultset generated by the command</returns>
		public static XmlReader ExecuteXmlReaderTypedParams(OleDbConnection connection, String spName, DataRow dataRow)
		{
			return new OleDb().ExecuteXmlReaderTypedParams( connection, spName, dataRow );
		}

		/// <summary>
		/// Execute a stored procedure via a OleDbCommand (that returns a resultset) against the specified OleDbTransaction 
		/// using the dataRow column values as the stored procedure's parameters values.
		/// This method will query the database to discover the parameters for the 
		/// stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
		/// </summary>
		/// <param name="transaction">A valid OleDbTransaction object</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="dataRow">The dataRow used to hold the stored procedure's parameter values.</param>
		/// <returns>An XmlReader containing the resultset generated by the command</returns>
		public static XmlReader ExecuteXmlReaderTypedParams(OleDbTransaction transaction, String spName, DataRow dataRow)
		{
			return new OleDb().ExecuteXmlReaderTypedParams( transaction, spName, dataRow );
		}
		#endregion

	}

	/// <summary>
	/// OleDbHelperParameterCache provides functions to leverage a static cache of procedure parameters, and the
	/// ability to discover parameters for stored procedures at run-time.
	/// </summary>
	public sealed class OleDbHelperParameterCache
	{
		#region private constructor

		//Since this class provides only static methods, make the default constructor private to prevent 
		//instances from being created with "new OleDbHelperParameterCache()"
		private OleDbHelperParameterCache() {}

		#endregion constructor

		#region caching functions

		/// <summary>
		/// Add parameter array to the cache
		/// </summary>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <param name="commandParameters">An array of OleDbParamters to be cached</param>
		public static void CacheParameterSet(string connectionString, string commandText, params OleDbParameter[] commandParameters)
		{
			new OleDb().CacheParameterSet(connectionString, commandText, commandParameters);
		}

		/// <summary>
		/// Retrieve a parameter array from the cache
		/// </summary>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="commandText">The stored procedure name or T-OleDb command</param>
		/// <returns>An array of OleDbParamters</returns>
		public static OleDbParameter[] GetCachedParameterSet(string connectionString, string commandText)
		{
			ArrayList tempValue = new ArrayList();
			IDataParameter[] OleDbP = new OleDb().GetCachedParameterSet(connectionString, commandText);
			foreach( IDataParameter parameter in OleDbP )
			{
				tempValue.Add( parameter );
			}
			return (OleDbParameter[])tempValue.ToArray( typeof(OleDbParameter) );
		}

		#endregion caching functions

		#region Parameter Discovery Functions

		/// <summary>
		/// Retrieves the set of OleDbParameters appropriate for the stored procedure
		/// </summary>
		/// <remarks>
		/// This method will query the database for this information, and then store it in a cache for future requests.
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <returns>An array of OleDbParameters</returns>
		public static OleDbParameter[] GetSpParameterSet(string connectionString, string spName)
		{
			ArrayList tempValue = new ArrayList();
			foreach( IDataParameter parameter in new OleDb().GetSpParameterSet( connectionString, spName ) )
			{
				tempValue.Add( parameter );
			}
			return (OleDbParameter[])tempValue.ToArray( typeof(OleDbParameter) );
		}

		/// <summary>
		/// Retrieves the set of OleDbParameters appropriate for the stored procedure
		/// </summary>
		/// <remarks>
		/// This method will query the database for this information, and then store it in a cache for future requests.
		/// </remarks>
		/// <param name="connectionString">A valid connection string for a OleDbConnection</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="includeReturnValueParameter">A bool value indicating whether the return value parameter should be included in the results</param>
		/// <returns>An array of OleDbParameters</returns>
		public static OleDbParameter[] GetSpParameterSet(string connectionString, string spName, bool includeReturnValueParameter)
		{
			ArrayList tempValue = new ArrayList();
			foreach( IDataParameter parameter in new OleDb().GetSpParameterSet( connectionString, spName, includeReturnValueParameter ) )
			{
				tempValue.Add( parameter );
			}
			return (OleDbParameter[])tempValue.ToArray( typeof(OleDbParameter) );
		}

		/// <summary>
		/// Retrieves the set of OleDbParameters appropriate for the stored procedure
		/// </summary>
		/// <remarks>
		/// This method will query the database for this information, and then store it in a cache for future requests.
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection object</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <returns>An array of OleDbParameters</returns>
		/// <exception cref="System.ArgumentNullException">Thrown if spName is null</exception>
		/// <exception cref="System.ArgumentNullException">Thrown if connection is null</exception>
		public static OleDbParameter[] GetSpParameterSet(IDbConnection connection, string spName)
		{
			return GetSpParameterSet(connection, spName, false);
		}

		/// <summary>
		/// Retrieves the set of OleDbParameters appropriate for the stored procedure
		/// </summary>
		/// <remarks>
		/// This method will query the database for this information, and then store it in a cache for future requests.
		/// </remarks>
		/// <param name="connection">A valid OleDbConnection object</param>
		/// <param name="spName">The name of the stored procedure</param>
		/// <param name="includeReturnValueParameter">A bool value indicating whether the return value parameter should be included in the results</param>
		/// <returns>An array of OleDbParameters</returns>
		/// <exception cref="System.ArgumentNullException">Thrown if spName is null</exception>
		/// <exception cref="System.ArgumentNullException">Thrown if connection is null</exception>
		public static OleDbParameter[] GetSpParameterSet(IDbConnection connection, string spName, bool includeReturnValueParameter)
		{
			ArrayList tempValue = new ArrayList();
			foreach( IDataParameter parameter in new OleDb().GetSpParameterSet( connection, spName, includeReturnValueParameter ) )
			{
				tempValue.Add( parameter );
			}
			return (OleDbParameter[])tempValue.ToArray( typeof(OleDbParameter) );
		}

		#endregion Parameter Discovery Functions

	}
}
