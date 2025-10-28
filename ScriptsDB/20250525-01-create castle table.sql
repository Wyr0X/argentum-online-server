CREATE TABLE IF NOT EXISTS "castle" (
    "id" integer NOT NULL,
    "guild_id"	integer NULL,
    "castle_type" integer NOT NULL,
    "state" integer NOT NULL DEFAULT 0,
    "timestamp" integer NOT NULL, 
    CONSTRAINT "fk_guild_castle" FOREIGN KEY("guild_id") REFERENCES "guilds"("id") ON DELETE CASCADE ON UPDATE CASCADE
    PRIMARY KEY ("id")
);