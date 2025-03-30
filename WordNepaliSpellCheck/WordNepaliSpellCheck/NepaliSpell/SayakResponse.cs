using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Web.Script.Serialization;

namespace WordNepaliSpellCheck.NepaliSpell        
{
    class SayakResponse
    {

        public String status { get; set; } = ResponseLiterals.DEFAULT_STATUS;

        public String messageId { get; set; } = TextUtility.EMPTY;

        public String message { get; set; } = TextUtility.EMPTY;

        public String type { get; set; } = ResponseLiterals.SAYAK_RESPONSE_TYPE;

        public String context { get; set; } = ResponseLiterals.SEMANTRO_CONTEXT;

        public String id { get; set; } = ResponseLiterals.SEMANTRO_SAYAK_RESPONSE_ID;

        /// <summary>
        /// Default constructor.
        /// </summary>
        public SayakResponse()
        {
            this.status = ResponseLiterals.DEFAULT_STATUS;
            this.messageId = TextUtility.EMPTY;
            this.message = TextUtility.EMPTY;
            this.type = ResponseLiterals.SAYAK_RESPONSE_TYPE;
            this.context = ResponseLiterals.SEMANTRO_CONTEXT;
            this.id = ResponseLiterals.SEMANTRO_SAYAK_RESPONSE_ID;
        }

        /// <summary>
        /// Overloaded constructor with dictionary parameters.
        /// </summary>
        /// <param name="responseData"></param>
        public SayakResponse(Dictionary<string, string> responseData)
        {
            this.status = responseData[ResponseLiterals.STATUS];
            this.messageId = responseData[ResponseLiterals.MESSAGE_ID];
            this.message = responseData[ResponseLiterals.MESSAGE];
            this.type = responseData[ResponseLiterals.TYPE];
            this.context = responseData[ResponseLiterals.CONTEXT];
            this.id = responseData[ResponseLiterals.ID];
        }
        
        public Boolean IsEmpty()
        {
            return this.message == TextUtility.EMPTY || this.messageId == TextUtility.EMPTY;
        } 
        
        public Boolean IsRedirectRegistration()
        {
            return this.messageId == ResponseLiterals.REDIRECT_TO_REGISTRATION;
        } 
        
        public Boolean IsRedirectServiceRenew()
        {
            return this.messageId == ResponseLiterals.REDIRECT_TO_SERVICE_RENEW;
        }
        
        public Boolean IsServiceNotAvailable()
        {
            return this.messageId == ResponseLiterals.SERVICE_NOT_AVAILABLE;
        } 

        public Boolean IsRedirectPayment()
        {
            return this.messageId == ResponseLiterals.REDIRECT_TO_SERVICE_PAYMENT;
        }

        public Boolean IsFailedStatus()
        {
            return this.status == ResponseLiterals.FALSE;
        }

        override
        public String ToString()
        {
            return new JavaScriptSerializer().Serialize(this);
        } 
        
        public static SayakResponse Empty()
        {            
            return new SayakResponse();
        }    

        public static SayakResponse buildWith(String jsonData)
        {
            try
            {
                return JsonConvert.DeserializeObject<SayakResponse>(jsonData.Replace(ResponseLiterals.ID_PLACE_HOLDER, TextUtility.EMPTY));
            }
            catch
            {
                return SayakResponse.Empty();
            }            
        }

        /// <summary>
        /// Factory to build the sayak response of offline, i.e. remote service is not available in the 
        /// current point of time.
        /// </summary>
        /// <returns></returns>
        public static SayakResponse serverNotAvailable()
        {
            var responseData = new Dictionary<String, String>() {
                { ResponseLiterals.STATUS, ResponseLiterals.DEFAULT_STATUS },
                { ResponseLiterals.MESSAGE_ID, ResponseLiterals.SERVICE_NOT_AVAILABLE },
                { ResponseLiterals.MESSAGE, ResponseLiterals.MESSAGE_SERVICE_NOT_AVAILABLE },
                { ResponseLiterals.TYPE, ResponseLiterals.SAYAK_RESPONSE_TYPE },
                { ResponseLiterals.CONTEXT, ResponseLiterals.SEMANTRO_CONTEXT },
                { ResponseLiterals.ID, ResponseLiterals.SEMANTRO_SAYAK_RESPONSE_ID }
            };
            return new SayakResponse(responseData);
        }

        /// <summary>
        /// Constructs remote message to display for the font service user.
        /// </summary>
        /// <returns>tupleOfURLMessage</returns>
        public Tuple<String, String> BuildRemoteURLMessage()
        {
            String remoteURL;
            String remoteMessage;
            if (this.IsRedirectRegistration())
            {
                remoteURL = RemoteIO.BuildServicePaymentURL();
                remoteMessage = $"Your device has not been registered yet for Nepali Spelling Service!\r\nGoto Sayak Registration Page: {remoteURL}";
            }
            else if (this.IsRedirectServiceRenew())
            {
                remoteURL = RemoteIO.BuildServiceRenewURL();
                remoteMessage = $"Nepali Spelling Service license expired!\r\nGoto Sayak Renew/Payment Page: {remoteURL}";
            }
            else if (this.IsRedirectPayment())
            {
                remoteURL = RemoteIO.BuildServiceRenewURL();
                remoteMessage = $"Nepali Spelling Service trial is over!\r\nGoto Sayak Renew/Payment Page: {remoteURL}";
            }
            else
            {
                remoteURL = RemoteIO.BuildContactURL();
                remoteMessage = $"Unable to make spelling suggestions.\r\nGoto Sayak contact page: {remoteURL} \r\n({this.message})";
            }
            return Tuple.Create(remoteURL, remoteMessage);
        }

    }

    class ResponseLiterals
    {      
        /// <summary>
        /// Message literals.
        /// </summary>                  
        public static readonly string TRUE = "true";
        public static readonly string FALSE = "false";
        public static readonly string DEFAULT_STATUS = FALSE;
        public static readonly string REDIRECT_TO_REGISTRATION = "REDIRECT_TO_REGISTRATION";
        public static readonly string REDIRECT_TO_SERVICE_RENEW = "REDIRECT_TO_SERVICE_RENEW";
        public static readonly string REDIRECT_TO_SERVICE_PAYMENT = "REDIRECT_TO_SERVICE_PAYMENT";
        public static readonly string SERVICE_NOT_AVAILABLE = "SERVICE_NOT_AVAILABLE";
        public static readonly string MESSAGE_TRIAL_OVER = "Trial Period Over!, We request you to register and make payment.";
        public static readonly string MESSAGE_SERVICE_EXPIRED = "Sorry, Your subscribed service duration is expired! We request you extend the spelling service.";
        public static readonly string MESSAGE_SERVICE_NOT_AVAILABLE = "Sorry, The remote spelling service is currently unavailable, please try after sometime. Exiting now ...";
        public static readonly string HEADING_SERVICE_NOT_AVAILABLE = "Service Not Available !";
        public static readonly string SAYAK_RESPONSE_TYPE = "SayakResponse";
        public static readonly string SEMANTRO_CONTEXT = "http://semantro.com";
        public static readonly string SEMANTRO_SAYAK_RESPONSE_ID = "http://semantro.com/SayakResponse";

        public static readonly string HEADING_EXISTING_SERVICE = "Existing Spell Service...";
        public static readonly string MESSAGE_EXISTING_SERVICE = "MS Word application has auto spell check enabled. Do you want to disable it for Nepali Language?";

        public static readonly string ID_PLACE_HOLDER = "@";

        /// <summary>
        /// Variable literals.
        /// </summary>
        public static readonly string STATUS = "status";
        public static readonly string MESSAGE = "message";
        public static readonly string MESSAGE_ID = "messageId";
        public static readonly string TYPE = "type";
        public static readonly string CONTEXT = "context";
        public static readonly string ID = "id";

        public static readonly string SAYAK_SERVICE_NAME = "Sayak Spelling Service";
    }
}
