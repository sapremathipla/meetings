import express, { Request, Response, NextFunction } from "express";
import cors from "cors";
import bodyParser from "body-parser";
import axios from "axios";
import { Pool } from "pg";
import Joi from "joi";
// import { Client } from "@microsoft/microsoft-graph-client";
// import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
// import { ClientSecretCredential } from "@azure/identity";
import { Sequelize, DataTypes, Model, Op } from "sequelize";

// Initialize Sequelize
const sequelize = new Sequelize("stage_sentinel_new", "postgres", "0xkmFCzp7RNarA0", {
  host: "staging-db-new.c69u4x9b0vhc.ap-south-1.rds.amazonaws.com",
  dialect: "postgres",
});
class Meeting extends Model {
  public id!: number;
  public purpose!: string;
  public startDateTime!: Date;
  public endDateTime!: Date;
  public attendees!: string;
  public room!: string;
  public host!: string;
  public tenantId!: string;
}
Meeting.init(
  {
    id: {
      type: DataTypes.INTEGER,
      autoIncrement: true,
      primaryKey: true,
    },
    purpose: {
      type: DataTypes.STRING,
      allowNull: false,
    },
    startDateTime: {
      type: DataTypes.DATE,
      allowNull: false,
    },
    endDateTime: {
      type: DataTypes.DATE,
      allowNull: false,
    },
    attendees: {
      type: DataTypes.TEXT,
      allowNull: false,
    },
    room: {
      type: DataTypes.STRING,
      allowNull: false,
    },
    host: {
      type: DataTypes.STRING,
      allowNull: false,
    },
    tenantId: {
      type: DataTypes.STRING,
      allowNull: false,
    },
  },
  {
    sequelize,
    modelName: "Meeting",
    tableName: "meeting",
    timestamps: true,
  }
);


const router = express.Router();
const app = express();
const port = 5000;
const clientId = "bc5b9a9f-8f73-48a9-80b6-6a8ce4d8e622";
const clientSecret = "2-Xx~gc0.lW~V61X471gjwiSpC6dgyzR3V";


const pool = new Pool({
  user: "postgres",
  host: "staging-db-new.c69u4x9b0vhc.ap-south-1.rds.amazonaws.com",
  database: "stage_sentinel_new",
  password: "0xkmFCzp7RNarA0",
  port: 5432,
});


app.use(express.json());
app.use(cors());
app.use(bodyParser.json());

const meetingSchema = Joi.object({
  purpose: Joi.string().required(),
  startDateTime: Joi.date().iso().required(),
  endDateTime: Joi.date().iso().required(),
  attendees: Joi.string().required(),
  room: Joi.string().required(),
  tenantId: Joi.string().required(),
  host: Joi.string().email().required(),
});


const getAccessToken = async (tenantId: string): Promise<string> => {
  const clientId = "bc5b9a9f-8f73-48a9-80b6-6a8ce4d8e622";
  const clientSecret = "2-Xx~gc0.lW~V61X471gjwiSpC6dgyzR3V";

  const response = await axios.post(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    }),
    { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
  );
  return response.data.access_token;
};

app.get("/mailservers", async (req: Request, res: Response) => {
  try {
    const result = await pool.query('SELECT * FROM "mailServer"');
    res.json(result.rows);
  } catch (error) {
    console.error("Error fetching mailservers:", error);
    res.status(500).send("Server error");
  }
});

app.post("/create-meeting", async (req: Request, res: Response): Promise<void> => {
  const { value } = meetingSchema.validate(req.body);
  const { purpose, startDateTime, endDateTime, attendees, room, tenantId, host } = value;
  req.app.locals.tenantId = tenantId;
  try {
    const accessToken = await getAccessToken(tenantId);
    const response = await axios.post(
      `https://graph.microsoft.com/v1.0/users/${host}/calendar/events`,
      {
        subject: purpose,
        start: {dateTime: startDateTime, timeZone: "UTC" },
        end: {dateTime: endDateTime, timeZone: "UTC" },
        location: {displayName: room },
        attendees: attendees.split(",").map((email: string) => ({
          emailAddress: {address:email.trim()},
          type: "required",
        })),
      },
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      }
    );
    res.status(200).json({ message: "Meeting created successfully", data: response.data });
  } catch (error) {
    if (axios.isAxiosError(error)) {
      console.error("Error creating event:", error.response?.data || error.message);
      res.status(500).json({ error: error.response?.data || "Failed to create event" });
    } else if (error instanceof Error) {
   
      console.error("Unexpected error:", error.message);
      res.status(500).json({ error: error.message || "Unexpected error occurred" });
    } else {
      console.error("Unknown error:", error);
      res.status(500).json({ error: "Unknown error occurred" });
    }
  }
});
//  function extractTenantId(req: Request, res: Response, next: NextFunction): void {
//   const { tenantId } = req.body;
//   if (!tenantId) {
//     res.status(400).json({ error: "tenantId is required" });
//     return;
//   }
//   req.app.locals.tenantId = tenantId;
//   next();
// }

// const createGraphClient = () => {
//   const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

//   const authProvider = new TokenCredentialAuthenticationProvider(credential, {
//     scopes: ["https://graph.microsoft.com/.default"],
//   });

//   return Client.initWithMiddleware({
//     authProvider,
//   });
// };
// router.get("/fetch-meetings",async(req:Request,res:Response)=>{
//   const{startDate,endDate,host}=req.query;

//   if (!startDate || !endDate || !host) {
//       res.status(400).json({
//       error: "startDate, endDate, and host parameters are required.",
//     });
//   }
//   try {
//     const graphClient = req.app.locals.graphClient;
//     const meetings = await graphClient.api(`/users/${host}/calendar/calendarView`)
//     .query({
//       startDateTime : new Date(startDate as string).toISOString(),
//       endDateTime : new Date(endDate as string).toISOString(),
//     }).header("prefer",'outlook.timezone="UTC"').get();
//     res.json(meetings.value);
//   } catch (error) {
//     console.log("Error fetching meetings",error);
//     res.status(500).json({
//       error :"failed to fetch meetings from outlook calendar",
//     });

//   }
// });
// export default router;

app.get("/meetings", async (req: Request, res: Response): Promise<void> => {
  const { date, host } = req.query;


  if (!date || !host) {
    res.status(400).json({ error: "Both 'date' and 'host' parameters are required." });
    return;
  }

  try {
    // Parse and validate the date
    const startDate = new Date(date as string);
    if (isNaN(startDate.getTime())) {
      res.status(400).json({ error: "Invalid date format." });
      return;
    }

    const endDate = new Date(startDate);
    endDate.setDate(startDate.getDate() + 1);

    const query = `
      SELECT m.*
      FROM meeting m
      INNER JOIN meetingParticipant mp ON m.id = mp.id
      WHERE mp.host = $1
        AND m.startTime >= $2
        AND m.endTime < $3
    `;
    const values = [host, startDate.toISOString(), endDate.toISOString()];
    const result = await pool.query(query, values);

    // Return the fetched meetings
    res.status(200).json(result.rows);
  } catch (error) {
    console.error("Error fetching meetings:", error);

    if (error instanceof Error) {
      res.status(500).json({ error: error.message });
    } else {
      res.status(500).json({ error: "Unknown error occurred." });
    }
  }
});


app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
});