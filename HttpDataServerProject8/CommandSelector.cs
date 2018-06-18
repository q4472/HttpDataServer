using Nskd;
using System;

namespace HttpDataServerProject8
{
    class CommandSelector
    {
        public static ResponsePackage Execute(RequestPackage rqp)
        {
            ResponsePackage rsp = null;

            if (rqp != null)
            {
                switch (rqp.Command)
                {
                    case "LoadAuction":
                        //Log.Write(String.Format("LoadAuction" + rqp.Command));
                        rsp = AuctionLoader.Load(rqp);
                        //Log.Write(String.Format(rsp.Status + rsp.Data.Tables[0].Rows[0][0].ToString()));
                        break;
                    case "LoadAuctionNumbers":
                        rsp = ZakupkiGovRu.LoadAuctionNumbers(rqp);
                        break;
                    case "WriteToConsole":
                        StoredProcedures.WriteRequestPackageToConsole(rqp);
                        break;
                    case "[Auctions].[dbo].[exists_auction_inf]":
                        rsp = Db.Exec(rqp);
                        break;
                    case "Prep.AddContractDirectory":
                        Prep.AddContractDirectory(rqp);
                        break;
                    case "Prep.PassToTender":
                        rsp = Prep.PassToTender(rqp);
                        break;
                    default:
                        break;
                }
            }
            return rsp;
        }
    }
}
